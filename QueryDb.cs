using System;
using System.Data.SQLite;

class Program
{
    static void Main()
    {
        string dbPath = @"C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Projects\Shared Test Model -00001\Filters\Shared_Test_Model_-00001_SleevePersistence.db";
        using (var conn = new SQLiteConnection($"Data Source={dbPath}"))
        {
            conn.Open();
            var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT ClashZoneId, SleeveInstanceId, ClusterSleeveInstanceId, IsClusterResolved, 
                                       HostOrientation, BoundingBoxMinY, BoundingBoxMaxY, BoundingBoxMinZ, BoundingBoxMaxZ 
                                FROM ClashZones WHERE MepElementCategory LIKE '%Accessor%'";
            var reader = cmd.ExecuteReader();
            Console.WriteLine("ClashZoneId | SlvId | ClsId | Res | Orient | Y-Range | Z-Range");
            Console.WriteLine("------------|-------|-------|-----|--------|---------|--------");
            while (reader.Read())
            {
                var zoneId = reader.GetString(0).Substring(0, 8);
                var slvId = reader.IsDBNull(1) ? -1 : reader.GetInt32(1);
                var clsId = reader.IsDBNull(2) ? -1 : reader.GetInt32(2);
                var resolved = reader.IsDBNull(3) ? 0 : reader.GetInt32(3);
                var orient = reader.IsDBNull(4) ? "?" : reader.GetString(4);
                var minY = reader.IsDBNull(5) ? 0 : Math.Round(reader.GetDouble(5), 2);
                var maxY = reader.IsDBNull(6) ? 0 : Math.Round(reader.GetDouble(6), 2);
                var minZ = reader.IsDBNull(7) ? 0 : Math.Round(reader.GetDouble(7), 2);
                var maxZ = reader.IsDBNull(8) ? 0 : Math.Round(reader.GetDouble(8), 2);
                Console.WriteLine($"{zoneId} | {slvId,5} | {clsId,5} | {resolved} | {orient,6} | [{minY},{maxY}] | [{minZ},{maxZ}]");
            }
        }
    }
}
