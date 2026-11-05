using System;
using System.IO;

namespace SEBatchReplace
{
    public enum RelationshipFailurePolicy
    {
        SkipAndContinue = 0,
        RollbackAndStop = 1
    }

    public sealed class ReplacementMapping
    {
        public string OldPartName { get; private set; }
        public string NewPartPath { get; private set; }

        public ReplacementMapping(string oldPartName, string newPartPath)
        {
            if (String.IsNullOrWhiteSpace(oldPartName))
                throw new ArgumentException("舊零件名稱不可空白。", "oldPartName");
            if (String.IsNullOrWhiteSpace(newPartPath))
                throw new ArgumentException("新零件路徑不可空白。", "newPartPath");

            OldPartName = Path.GetFileName(oldPartName.Trim());
            if (!OldPartName.EndsWith(".par", StringComparison.OrdinalIgnoreCase))
                OldPartName += ".par";

            NewPartPath = newPartPath;
        }

        public override string ToString()
        {
            return OldPartName + "  →  " + Path.GetFileName(NewPartPath);
        }
    }
}
