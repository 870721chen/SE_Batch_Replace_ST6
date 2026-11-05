using System;

namespace SEBatchReplace
{
    public enum RelationshipFailurePolicy
    {
        SkipAndContinue = 0,
        RollbackAndStop = 1
    }

    public sealed class ReplacementMapping
    {
        public string OldPartPath { get; private set; }
        public string NewPartPath { get; private set; }

        public ReplacementMapping(string oldPartPath, string newPartPath)
        {
            if (String.IsNullOrWhiteSpace(oldPartPath))
                throw new ArgumentException("舊零件路徑不可空白。", "oldPartPath");
            if (String.IsNullOrWhiteSpace(newPartPath))
                throw new ArgumentException("新零件路徑不可空白。", "newPartPath");

            OldPartPath = oldPartPath;
            NewPartPath = newPartPath;
        }

        public override string ToString()
        {
            return System.IO.Path.GetFileName(OldPartPath) + "  →  " +
                   System.IO.Path.GetFileName(NewPartPath);
        }
    }
}
