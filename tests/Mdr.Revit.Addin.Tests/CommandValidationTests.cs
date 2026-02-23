using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Addin.Commands;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class CommandValidationTests
    {
        [Fact]
        public async Task LoginCommand_WithoutUsername_Throws()
        {
            LoginCommand command = new LoginCommand();

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                command.ExecuteAsync(
                    new LoginCommandRequest
                    {
                        BaseUrl = "http://127.0.0.1:8000",
                        Username = "",
                        Password = "secret",
                    },
                    CancellationToken.None));
        }

        [Fact]
        public async Task PublishSheetsCommand_WithoutProjectCode_Throws()
        {
            PublishSheetsCommand command = new PublishSheetsCommand();

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                command.ExecuteAsync(
                    new PublishSheetsCommandRequest
                    {
                        BaseUrl = "http://127.0.0.1:8000",
                        Username = "u",
                        Password = "p",
                        ProjectCode = "",
                    },
                    CancellationToken.None));
        }

        [Fact]
        public async Task PushSchedulesCommand_WithoutProfileCode_Throws()
        {
            PushSchedulesCommand command = new PushSchedulesCommand();

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                command.ExecuteAsync(
                    new PushSchedulesCommandRequest
                    {
                        BaseUrl = "http://127.0.0.1:8000",
                        Username = "u",
                        Password = "p",
                        ProjectCode = "PRJ-001",
                        ProfileCode = "",
                    },
                    CancellationToken.None));
        }

        [Fact]
        public async Task SyncSiteLogsCommand_WithoutModelGuid_Throws()
        {
            SyncSiteLogsCommand command = new SyncSiteLogsCommand();

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                command.ExecuteAsync(
                    new SyncSiteLogsCommandRequest
                    {
                        BaseUrl = "http://127.0.0.1:8000",
                        Username = "u",
                        Password = "p",
                        ProjectCode = "PRJ-001",
                        ClientModelGuid = "",
                    },
                    CancellationToken.None));
        }

        [Fact]
        public async Task GoogleSyncCommand_WithoutSpreadsheetId_Throws()
        {
            GoogleSyncCommand command = new GoogleSyncCommand();

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                command.ExecuteAsync(
                    new GoogleSyncCommandRequest
                    {
                        Direction = "export",
                        SpreadsheetId = "",
                        WorksheetName = "Sheet1",
                        GoogleClientId = "client",
                        GoogleClientSecret = "secret",
                    },
                    CancellationToken.None));
        }

        [Fact]
        public async Task CheckUpdatesCommand_WithoutRepo_Throws()
        {
            CheckUpdatesCommand command = new CheckUpdatesCommand();

            await Assert.ThrowsAsync<InvalidOperationException>(() =>
                command.ExecuteAsync(
                    new CheckUpdatesCommandRequest
                    {
                        CurrentVersion = "1.0.0",
                        GithubRepo = "",
                    },
                    CancellationToken.None));
        }
    }
}
