using System;

namespace Mdr.Revit.RevitAdapter.Helpers
{
    public sealed class RevitTransactionRunner
    {
        public void Run(string transactionName, Action action)
        {
            if (string.IsNullOrWhiteSpace(transactionName))
            {
                throw new ArgumentException("Transaction name is required.", nameof(transactionName));
            }

            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            // Transaction wrapper will call Autodesk.Revit.DB.Transaction in real implementation.
            action();
        }
    }
}
