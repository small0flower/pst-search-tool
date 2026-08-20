// ─────────────────────────────────────────────────────────────
// 僅供 Linux/mono 編譯驗證用的 System.Data.SQLite 最小 stub。
// 不屬於專案本體；真實建置使用 NuGet 的 System.Data.SQLite.Core。
// ─────────────────────────────────────────────────────────────
namespace System.Data.SQLite
{
    public enum SQLiteJournalModeEnum { Default, Delete, Persist, Off, Truncate, Memory, Wal }

    public class SQLiteConnectionStringBuilder
    {
        public string DataSource { get; set; }
        public int Version { get; set; }
        public SQLiteJournalModeEnum JournalMode { get; set; }
        public bool Pooling { get; set; }
        public override string ToString() { return "Data Source=" + DataSource; }
    }

    public class SQLiteConnection : IDisposable
    {
        public SQLiteConnection(string connectionString) { }
        public void Open() { }
        public SQLiteCommand CreateCommand() { return new SQLiteCommand(); }
        public SQLiteTransaction BeginTransaction() { return new SQLiteTransaction(); }
        public long LastInsertRowId { get { return 0; } }
        public void Close() { }
        public void Dispose() { }
    }

    public class SQLiteTransaction : IDisposable
    {
        public void Commit() { }
        public void Rollback() { }
        public void Dispose() { }
    }

    public class SQLiteParameter
    {
        public SQLiteParameter(string name, object value) { ParameterName = name; Value = value; }
        public string ParameterName { get; set; }
        public object Value { get; set; }
    }

    public class SQLiteParameterCollection
    {
        public void AddWithValue(string name, object value) { }
        public void AddRange(SQLiteParameter[] pars) { }
        public void Clear() { }
    }

    public class SQLiteCommand : IDisposable
    {
        public string CommandText { get; set; }
        public SQLiteTransaction Transaction { get; set; }
        private readonly SQLiteParameterCollection _p = new SQLiteParameterCollection();
        public SQLiteParameterCollection Parameters { get { return _p; } }
        public int ExecuteNonQuery() { return 0; }
        public object ExecuteScalar() { return null; }
        public SQLiteDataReader ExecuteReader() { return new SQLiteDataReader(); }
        public void Dispose() { }
    }

    public class SQLiteDataReader : IDisposable
    {
        public bool Read() { return false; }
        public long GetInt64(int i) { return 0; }
        public string GetString(int i) { return ""; }
        public bool IsDBNull(int i) { return true; }
        public void Dispose() { }
    }
}
