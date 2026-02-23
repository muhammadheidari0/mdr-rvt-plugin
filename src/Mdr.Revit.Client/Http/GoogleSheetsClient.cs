using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Auth;
using Mdr.Revit.Client.Retry;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Client.Http
{
    public sealed class GoogleSheetsClient : IGoogleSheetsClient, IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly IGoogleTokenProvider _tokenProvider;
        private readonly RetryPolicy _retryPolicy;
        private bool _disposed;

        public GoogleSheetsClient(IGoogleTokenProvider tokenProvider)
            : this(
                new HttpClient { BaseAddress = new Uri("https://sheets.googleapis.com") },
                tokenProvider,
                new RetryPolicy())
        {
        }

        public GoogleSheetsClient(
            HttpClient httpClient,
            IGoogleTokenProvider tokenProvider,
            RetryPolicy retryPolicy)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _tokenProvider = tokenProvider ?? throw new ArgumentNullException(nameof(tokenProvider));
            _retryPolicy = retryPolicy ?? throw new ArgumentNullException(nameof(retryPolicy));
        }

        public Task<GoogleSheetReadResult> ReadRowsAsync(
            GoogleSheetSyncProfile profile,
            CancellationToken cancellationToken)
        {
            if (profile == null)
            {
                throw new ArgumentNullException(nameof(profile));
            }

            return _retryPolicy.ExecuteAsync(async token =>
            {
                EnsureNotDisposed();
                string range = Uri.EscapeDataString(profile.WorksheetName + "!A:ZZ");
                string path = "/v4/spreadsheets/" + Uri.EscapeDataString(profile.SpreadsheetId) + "/values/" + range;

                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, path))
                {
                    await AttachAuthorizationAsync(request, token).ConfigureAwait(false);
                    using (HttpResponseMessage response = await _httpClient.SendAsync(request, token).ConfigureAwait(false))
                    {
                        string body = response.Content == null
                            ? string.Empty
                            : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        if (!response.IsSuccessStatusCode)
                        {
                            throw new HttpRequestException(
                                "Google Sheets read failed with status " + (int)response.StatusCode + ": " + body);
                        }

                        return ParseReadResponse(body, profile.AnchorColumn);
                    }
                }
            }, cancellationToken);
        }

        public Task<GoogleSheetWriteResult> WriteRowsAsync(
            GoogleSheetSyncProfile profile,
            IReadOnlyList<ScheduleSyncRow> rows,
            CancellationToken cancellationToken)
        {
            if (profile == null)
            {
                throw new ArgumentNullException(nameof(profile));
            }

            rows ??= Array.Empty<ScheduleSyncRow>();

            return _retryPolicy.ExecuteAsync(async token =>
            {
                EnsureNotDisposed();

                IReadOnlyList<string> headers = ResolveHeaders(profile, rows);
                List<List<string>> matrix = BuildWriteMatrix(headers, rows);
                string escapedSheet = Uri.EscapeDataString(profile.SpreadsheetId);
                string escapedRange = Uri.EscapeDataString(profile.WorksheetName + "!A:ZZ");

                using (HttpRequestMessage clear = new HttpRequestMessage(
                    HttpMethod.Post,
                    "/v4/spreadsheets/" + escapedSheet + "/values/" + escapedRange + ":clear"))
                {
                    await AttachAuthorizationAsync(clear, token).ConfigureAwait(false);
                    clear.Content = new StringContent("{}", Encoding.UTF8, "application/json");
                    using (HttpResponseMessage ignored = await _httpClient.SendAsync(clear, token).ConfigureAwait(false))
                    {
                        _ = ignored;
                    }
                }

                string updateRange = Uri.EscapeDataString(profile.WorksheetName + "!A1");
                string path = "/v4/spreadsheets/" + escapedSheet + "/values/" + updateRange + "?valueInputOption=RAW";
                string payload = JsonSerializer.Serialize(
                    new
                    {
                        range = profile.WorksheetName + "!A1",
                        majorDimension = "ROWS",
                        values = matrix,
                    });

                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Put, path))
                {
                    await AttachAuthorizationAsync(request, token).ConfigureAwait(false);
                    request.Content = new StringContent(payload, Encoding.UTF8, "application/json");
                    using (HttpResponseMessage response = await _httpClient.SendAsync(request, token).ConfigureAwait(false))
                    {
                        string body = response.Content == null
                            ? string.Empty
                            : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        if (!response.IsSuccessStatusCode)
                        {
                            throw new HttpRequestException(
                                "Google Sheets write failed with status " + (int)response.StatusCode + ": " + body);
                        }

                        return ParseWriteResponse(body, rows.Count);
                    }
                }
            }, cancellationToken);
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _httpClient.Dispose();
            if (_tokenProvider is IDisposable disposable)
            {
                disposable.Dispose();
            }

            _disposed = true;
        }

        private static GoogleSheetReadResult ParseReadResponse(string json, string anchorColumn)
        {
            GoogleSheetReadResult result = new GoogleSheetReadResult();
            if (string.IsNullOrWhiteSpace(json))
            {
                return result;
            }

            using (JsonDocument document = JsonDocument.Parse(json))
            {
                if (!document.RootElement.TryGetProperty("values", out JsonElement valuesElement) ||
                    valuesElement.ValueKind != JsonValueKind.Array)
                {
                    return result;
                }

                List<List<string>> rows = new List<List<string>>();
                foreach (JsonElement row in valuesElement.EnumerateArray())
                {
                    if (row.ValueKind != JsonValueKind.Array)
                    {
                        continue;
                    }

                    List<string> columns = new List<string>();
                    foreach (JsonElement cell in row.EnumerateArray())
                    {
                        columns.Add(cell.GetString() ?? string.Empty);
                    }

                    rows.Add(columns);
                }

                if (rows.Count == 0)
                {
                    return result;
                }

                IReadOnlyList<string> headers = rows[0];
                result.Headers = headers.ToArray();

                int anchorIndex = FindHeaderIndex(headers, anchorColumn);
                HashSet<string> seenAnchors = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++)
                {
                    IReadOnlyList<string> source = rows[rowIndex];
                    ScheduleSyncRow parsed = new ScheduleSyncRow();
                    for (int colIndex = 0; colIndex < headers.Count; colIndex++)
                    {
                        string header = headers[colIndex] ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(header))
                        {
                            continue;
                        }

                        string value = colIndex < source.Count ? source[colIndex] ?? string.Empty : string.Empty;
                        parsed.Cells[header] = value;
                    }

                    if (anchorIndex < 0 || anchorIndex >= source.Count || string.IsNullOrWhiteSpace(source[anchorIndex]))
                    {
                        parsed.ChangeState = ScheduleSyncStates.Error;
                        parsed.Errors.Add(new ScheduleSyncError
                        {
                            Code = "anchor_missing",
                            Message = "Anchor column value is missing.",
                        });
                    }
                    else
                    {
                        parsed.AnchorUniqueId = source[anchorIndex];
                        if (!seenAnchors.Add(parsed.AnchorUniqueId))
                        {
                            parsed.ChangeState = ScheduleSyncStates.Error;
                            parsed.Errors.Add(new ScheduleSyncError
                            {
                                Code = "anchor_duplicate",
                                Message = "Anchor value is duplicated in Google Sheet.",
                            });
                        }
                    }

                    result.Rows.Add(parsed);
                }
            }

            return result;
        }

        private static GoogleSheetWriteResult ParseWriteResponse(string json, int defaultRows)
        {
            GoogleSheetWriteResult result = new GoogleSheetWriteResult
            {
                UpdatedRows = defaultRows,
            };

            if (string.IsNullOrWhiteSpace(json))
            {
                return result;
            }

            using (JsonDocument document = JsonDocument.Parse(json))
            {
                if (document.RootElement.TryGetProperty("updatedRows", out JsonElement updatedRowsElement) &&
                    updatedRowsElement.TryGetInt32(out int updatedRows))
                {
                    result.UpdatedRows = Math.Max(0, updatedRows - 1);
                }

                if (document.RootElement.TryGetProperty("updatedRange", out JsonElement updatedRangeElement))
                {
                    result.UpdatedRange = updatedRangeElement.GetString() ?? string.Empty;
                }
            }

            return result;
        }

        private static IReadOnlyList<string> ResolveHeaders(
            GoogleSheetSyncProfile profile,
            IReadOnlyList<ScheduleSyncRow> rows)
        {
            List<string> headers = new List<string>();

            AddHeader(headers, profile.AnchorColumn);
            AddHeader(headers, "MDR_ELEMENT_ID");
            for (int i = 0; i < profile.ColumnMappings.Count; i++)
            {
                AddHeader(headers, profile.ColumnMappings[i].SheetColumn);
            }

            for (int i = 0; i < rows.Count; i++)
            {
                foreach (KeyValuePair<string, string> pair in rows[i].Cells)
                {
                    AddHeader(headers, pair.Key);
                }
            }

            return headers;
        }

        private static List<List<string>> BuildWriteMatrix(
            IReadOnlyList<string> headers,
            IReadOnlyList<ScheduleSyncRow> rows)
        {
            List<List<string>> matrix = new List<List<string>>(rows.Count + 1)
            {
                headers.ToList(),
            };

            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                ScheduleSyncRow row = rows[rowIndex];
                List<string> values = new List<string>(headers.Count);
                for (int headerIndex = 0; headerIndex < headers.Count; headerIndex++)
                {
                    string header = headers[headerIndex];
                    if (header.Equals("MDR_UNIQUE_ID", StringComparison.OrdinalIgnoreCase))
                    {
                        values.Add(row.AnchorUniqueId ?? string.Empty);
                        continue;
                    }

                    if (header.Equals("MDR_ELEMENT_ID", StringComparison.OrdinalIgnoreCase))
                    {
                        values.Add(row.ElementId ?? string.Empty);
                        continue;
                    }

                    if (!row.Cells.TryGetValue(header, out string value))
                    {
                        value = string.Empty;
                    }

                    values.Add(value ?? string.Empty);
                }

                matrix.Add(values);
            }

            return matrix;
        }

        private static void AddHeader(List<string> headers, string value)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }

            if (headers.Any(x => x.Equals(normalized, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            headers.Add(normalized);
        }

        private static int FindHeaderIndex(IReadOnlyList<string> headers, string expected)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                if (string.Equals(headers[i], expected, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private async Task AttachAuthorizationAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            string token = await _tokenProvider.GetAccessTokenAsync(cancellationToken).ConfigureAwait(false);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
        }

        private void EnsureNotDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(GoogleSheetsClient));
            }
        }
    }
}
