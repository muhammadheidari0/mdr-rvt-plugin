using System;
using System.Collections.Generic;
using System.IO;

namespace Mdr.Revit.Core.Models
{
    public sealed class PublishBatchRequest
    {
        public string RunClientId { get; set; } = Guid.NewGuid().ToString("N");

        public string ProjectCode { get; set; } = string.Empty;

        public string RevitVersion { get; set; } = "2026";

        public string ModelGuid { get; set; } = string.Empty;

        public string ModelTitle { get; set; } = string.Empty;

        public string PluginVersion { get; set; } = string.Empty;

        public List<PublishSheetItem> Items { get; } = new List<PublishSheetItem>();

        public List<PublishFileManifestItem> FilesManifest { get; } = new List<PublishFileManifestItem>();

        public bool HasAnyLocalFile()
        {
            foreach (PublishSheetItem item in Items)
            {
                if (item == null)
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(item.PdfFilePath) || !string.IsNullOrWhiteSpace(item.NativeFilePath))
                {
                    return true;
                }
            }

            return false;
        }

        public IReadOnlyList<PublishFileManifestItem> BuildFilesManifest()
        {
            FilesManifest.Clear();

            foreach (PublishSheetItem item in Items)
            {
                if (item == null)
                {
                    continue;
                }

                string pdfFileName = item.GetPdfFileName();
                string nativeFileName = item.GetNativeFileName();

                if (string.IsNullOrWhiteSpace(pdfFileName) && string.IsNullOrWhiteSpace(nativeFileName))
                {
                    continue;
                }

                FilesManifest.Add(new PublishFileManifestItem
                {
                    ItemIndex = item.ItemIndex,
                    SheetUniqueId = item.SheetUniqueId,
                    PdfFileName = pdfFileName,
                    NativeFileName = nativeFileName,
                    FileSha256 = item.FileSha256,
                });
            }

            return FilesManifest;
        }

        public PublishBatchRequest CreateRetryRequest(PublishBatchResponse response)
        {
            if (response == null)
            {
                throw new ArgumentNullException(nameof(response));
            }

            PublishBatchRequest retry = new PublishBatchRequest
            {
                RunClientId = Guid.NewGuid().ToString("N"),
                ProjectCode = ProjectCode,
                RevitVersion = RevitVersion,
                ModelGuid = ModelGuid,
                ModelTitle = ModelTitle,
                PluginVersion = PluginVersion,
            };

            HashSet<int> failedIndexes = new HashSet<int>();
            foreach (PublishItemResult item in response.Items)
            {
                if (!item.IsFailed())
                {
                    continue;
                }

                failedIndexes.Add(item.ItemIndex);
            }

            foreach (PublishSheetItem item in Items)
            {
                if (failedIndexes.Contains(item.ItemIndex))
                {
                    retry.Items.Add(item.Clone());
                }
            }

            retry.BuildFilesManifest();
            return retry;
        }
    }

    public sealed class PublishSheetItem
    {
        public int ItemIndex { get; set; }

        public string SheetUniqueId { get; set; } = string.Empty;

        public string SheetNumber { get; set; } = string.Empty;

        public string SheetName { get; set; } = string.Empty;

        public string DocNumber { get; set; } = string.Empty;

        public string RequestedRevision { get; set; } = string.Empty;

        public string StatusCode { get; set; } = string.Empty;

        public bool IncludeNative { get; set; }

        public DocumentMetadata Metadata { get; set; } = new DocumentMetadata();

        public string PdfFilePath { get; set; } = string.Empty;

        public string NativeFilePath { get; set; } = string.Empty;

        public string FileSha256 { get; set; } = string.Empty;

        public string GetPdfFileName()
        {
            return TryGetFileName(PdfFilePath);
        }

        public string GetNativeFileName()
        {
            return TryGetFileName(NativeFilePath);
        }

        public PublishSheetItem Clone()
        {
            return new PublishSheetItem
            {
                ItemIndex = ItemIndex,
                SheetUniqueId = SheetUniqueId,
                SheetNumber = SheetNumber,
                SheetName = SheetName,
                DocNumber = DocNumber,
                RequestedRevision = RequestedRevision,
                StatusCode = StatusCode,
                IncludeNative = IncludeNative,
                Metadata = Metadata?.Clone() ?? new DocumentMetadata(),
                PdfFilePath = PdfFilePath,
                NativeFilePath = NativeFilePath,
                FileSha256 = FileSha256,
            };
        }

        private static string TryGetFileName(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFileName(path.Trim()) ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }

    public sealed class DocumentMetadata
    {
        public string MdrCode { get; set; } = string.Empty;

        public string Phase { get; set; } = string.Empty;

        public string Discipline { get; set; } = string.Empty;

        public string Package { get; set; } = string.Empty;

        public string Block { get; set; } = string.Empty;

        public string Level { get; set; } = string.Empty;

        public string Subject { get; set; } = string.Empty;

        public DocumentMetadata Clone()
        {
            return new DocumentMetadata
            {
                MdrCode = MdrCode,
                Phase = Phase,
                Discipline = Discipline,
                Package = Package,
                Block = Block,
                Level = Level,
                Subject = Subject,
            };
        }
    }

    public sealed class PublishFileManifestItem
    {
        public int ItemIndex { get; set; }

        public string SheetUniqueId { get; set; } = string.Empty;

        public string PdfFileName { get; set; } = string.Empty;

        public string NativeFileName { get; set; } = string.Empty;

        public string FileSha256 { get; set; } = string.Empty;
    }

    public sealed class PublishBatchResponse
    {
        public string RunId { get; set; } = string.Empty;

        public PublishBatchSummary Summary { get; set; } = new PublishBatchSummary();

        public List<PublishItemResult> Items { get; } = new List<PublishItemResult>();
    }

    public sealed class PublishBatchSummary
    {
        public int RequestedCount { get; set; }

        public int SuccessCount { get; set; }

        public int FailedCount { get; set; }

        public int DuplicateCount { get; set; }

        public string Status { get; set; } = string.Empty;
    }

    public sealed class PublishItemResult
    {
        public int ItemIndex { get; set; }

        public string State { get; set; } = string.Empty;

        public long? DocumentId { get; set; }

        public string DocNumber { get; set; } = string.Empty;

        public string AppliedRevision { get; set; } = string.Empty;

        public long? PdfFileId { get; set; }

        public long? NativeFileId { get; set; }

        public string ErrorCode { get; set; } = string.Empty;

        public string ErrorMessage { get; set; } = string.Empty;

        public bool IsFailed()
        {
            return string.Equals(State, "failed", StringComparison.OrdinalIgnoreCase);
        }
    }
}
