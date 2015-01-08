// Guids.cs
// MUST match guids.h
using System;

namespace chrisbjohnson.TaskListPlus
{
    static class GuidList
    {
        public const string guidTaskListPlusPkgString = "8a76dea0-8591-4668-abc5-c7dacc31369d";
        public const string guidTaskListPlusCmdSetString = "8678c7ce-b88f-4d8d-be72-59963bfcda7e";
        public const string guidToolWindowPersistanceString = "c8bc3106-258b-4638-984c-db2763a9ca9b";

        public static readonly Guid guidTaskListPlusCmdSet = new Guid(guidTaskListPlusCmdSetString);
    };
}