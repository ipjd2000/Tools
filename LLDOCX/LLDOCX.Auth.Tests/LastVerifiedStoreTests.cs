using LLDOCX.Auth;

namespace LLDOCX.Auth.Tests;

[TestClass]
public class LastVerifiedStoreTests
{
    string _dir;
    string _file;

    [TestInitialize]
    public void Init()
    {
        _dir = Path.Combine(Path.GetTempPath(), "lldocx-auth-tests-" + Guid.NewGuid().ToString("N"));
        _file = Path.Combine(_dir, "gate.dat");
    }

    [TestCleanup]
    public void Cleanup()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, true);
    }

    [TestMethod]
    public void Round_trips_a_record_and_creates_the_folder()
    {
        var store = new DpapiLastVerifiedStore(_file);
        var when = new DateTime(2026, 11, 1, 8, 30, 0, DateTimeKind.Utc);
        store.Save(new VerifiedRecord(when, "acct-1", "pat@lozaip.com"));
        var loaded = new DpapiLastVerifiedStore(_file).Load();
        Assert.AreEqual(when, loaded.VerifiedUtc);
        Assert.AreEqual(DateTimeKind.Utc, loaded.VerifiedUtc.Kind);
        Assert.AreEqual("acct-1", loaded.AccountId);
        Assert.AreEqual("pat@lozaip.com", loaded.Username);
    }

    [TestMethod]
    public void File_is_not_plain_text()
    {
        new DpapiLastVerifiedStore(_file).Save(new VerifiedRecord(DateTime.UtcNow, "acct-1", "pat@lozaip.com"));
        StringAssert.DoesNotMatch(File.ReadAllText(_file), new System.Text.RegularExpressions.Regex("lozaip"));
    }

    [TestMethod]
    public void Missing_file_loads_as_null()
    {
        Assert.IsNull(new DpapiLastVerifiedStore(_file).Load());
    }

    [TestMethod]
    public void Corrupt_file_loads_as_null()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllBytes(_file, new byte[] { 1, 2, 3, 4, 5 });
        Assert.IsNull(new DpapiLastVerifiedStore(_file).Load());
    }

    [TestMethod]
    public void Clear_deletes_the_record()
    {
        var store = new DpapiLastVerifiedStore(_file);
        store.Save(new VerifiedRecord(DateTime.UtcNow, "a", "u"));
        store.Clear();
        Assert.IsNull(store.Load());
        store.Clear(); // clearing twice is harmless
    }
}
