using System;
using System.IO;
using System.Linq;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Validation
{
    public static class BusinessRules
    {
        public static void EnsurePublishRequestIsValid(PublishBatchRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            EnsureProjectCode(request.ProjectCode);

            if (request.Items.Count == 0)
            {
                throw new InvalidOperationException("At least one sheet item is required for publish.");
            }

            bool hasInvalidSheet = request.Items.Any(x => string.IsNullOrWhiteSpace(x.SheetUniqueId));
            if (hasInvalidSheet)
            {
                throw new InvalidOperationException("Every publish item must have SheetUniqueId.");
            }

            bool hasMissingRevision = request.Items.Any(x => string.IsNullOrWhiteSpace(x.RequestedRevision));
            if (hasMissingRevision)
            {
                throw new InvalidOperationException("Every publish item must have RequestedRevision.");
            }

            bool hasNoPdfOrHash = request.Items.Any(x =>
                string.IsNullOrWhiteSpace(x.PdfFilePath) &&
                string.IsNullOrWhiteSpace(x.FileSha256));
            if (hasNoPdfOrHash)
            {
                throw new InvalidOperationException("Every publish item must include PdfFilePath or FileSha256.");
            }

            bool hasMissingPdfFile = request.Items.Any(x =>
                !string.IsNullOrWhiteSpace(x.PdfFilePath) &&
                !File.Exists(x.PdfFilePath));
            if (hasMissingPdfFile)
            {
                throw new InvalidOperationException("At least one PdfFilePath does not exist.");
            }

            bool hasMissingNativeFile = request.Items.Any(x =>
                !string.IsNullOrWhiteSpace(x.NativeFilePath) &&
                !File.Exists(x.NativeFilePath));
            if (hasMissingNativeFile)
            {
                throw new InvalidOperationException("At least one NativeFilePath does not exist.");
            }
        }

        public static void EnsureScheduleRequestIsValid(ScheduleIngestRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            EnsureProjectCode(request.ProjectCode);

            if (string.IsNullOrWhiteSpace(request.ProfileCode))
            {
                throw new InvalidOperationException("Schedule ProfileCode is required.");
            }

            if (request.Rows.Count == 0)
            {
                throw new InvalidOperationException("At least one schedule row is required.");
            }
        }

        public static void EnsureManifestRequestIsValid(SiteLogManifestRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            EnsureProjectCode(request.ProjectCode);

            if (request.Limit <= 0 || request.Limit > 10000)
            {
                throw new InvalidOperationException("Manifest limit must be between 1 and 10000.");
            }
        }

        private static void EnsureProjectCode(string projectCode)
        {
            if (string.IsNullOrWhiteSpace(projectCode))
            {
                throw new InvalidOperationException("ProjectCode is required.");
            }
        }
    }
}
