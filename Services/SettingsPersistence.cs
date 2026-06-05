using System.Text;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using Windows.Storage;
using wows_ime.Views;

namespace wows_ime.Services;

internal sealed record PersistedGamePath(string DisplayName, string Path);

internal sealed record PersistedInputMethod(string DisplayName, string Category);

internal static class SettingsPersistence
{
    private const int CurrentSchemaVersion = 1;
    private const string SchemaVersionKey = "Settings.SchemaVersion";
    private const string SelectedGamePathKey = "Game.SelectedPath";
    private const string LegacyConfigFileName = "config.json";
    private const string MigratedLegacyConfigFileName = "config.json.migrated";
    private const string DatabaseFileName = "settings.db";

    internal static void Initialize()
    {
        try
        {
            ApplicationData.Current.LocalSettings.Values[SchemaVersionKey] = CurrentSchemaVersion;
            using var connection = OpenConnection();
            ExecuteNonQuery(connection, """
                CREATE TABLE IF NOT EXISTS custom_game_paths (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    display_name TEXT NOT NULL,
                    path TEXT NOT NULL UNIQUE,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                """);

            ExecuteNonQuery(connection, """
                CREATE TABLE IF NOT EXISTS custom_input_methods (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    display_name TEXT NOT NULL UNIQUE,
                    category TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                """);

            MigrateLegacyConfigJsonIfNeeded(connection);
        }
        catch
        {
            // Keep the app usable even when the local database cannot be initialized.
        }
    }

    internal static string? LoadSelectedGamePath()
    {
        try
        {
            return ApplicationData.Current.LocalSettings.Values.TryGetValue(SelectedGamePathKey, out var value)
                ? value as string
                : null;
        }
        catch
        {
            return null;
        }
    }

    internal static void SaveSelectedGamePath(string selectedGamePath)
    {
        try
        {
            ApplicationData.Current.LocalSettings.Values[SelectedGamePathKey] = selectedGamePath;
        }
        catch
        {
            // Keep failures silent to avoid breaking the main workflow.
        }
    }

    internal static List<PersistedGamePath> LoadCustomGamePaths()
    {
        var results = new List<PersistedGamePath>();
        try
        {
            using var connection = OpenConnection();
            using var command = connection.CreateCommand();
            command.CommandText = """
                SELECT display_name, path
                FROM custom_game_paths
                ORDER BY id;
                """;

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                results.Add(new PersistedGamePath(reader.GetString(0), reader.GetString(1)));
            }
        }
        catch
        {
            return results;
        }

        return results;
    }

    internal static void SaveCustomGamePaths(IEnumerable<PersistedGamePath> customGamePaths)
    {
        try
        {
            using var connection = OpenConnection();
            using var transaction = connection.BeginTransaction();
            ExecuteNonQuery(connection, transaction, "DELETE FROM custom_game_paths;");

            foreach (var path in customGamePaths)
            {
                using var command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = """
                    INSERT INTO custom_game_paths (display_name, path, updated_at)
                    VALUES ($displayName, $path, CURRENT_TIMESTAMP);
                    """;
                command.Parameters.AddWithValue("$displayName", path.DisplayName);
                command.Parameters.AddWithValue("$path", path.Path);
                command.ExecuteNonQuery();
            }

            transaction.Commit();
        }
        catch
        {
            // Keep failures silent to avoid breaking the main workflow.
        }
    }

    internal static List<PersistedInputMethod> LoadCustomInputMethods()
    {
        var results = new List<PersistedInputMethod>();
        try
        {
            using var connection = OpenConnection();
            using var command = connection.CreateCommand();
            command.CommandText = """
                SELECT display_name, category
                FROM custom_input_methods
                ORDER BY id;
                """;

            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                results.Add(new PersistedInputMethod(reader.GetString(0), reader.GetString(1)));
            }
        }
        catch
        {
            return results;
        }

        return results;
    }

    internal static void SaveCustomInputMethods(IEnumerable<PersistedInputMethod> customInputMethods)
    {
        try
        {
            using var connection = OpenConnection();
            using var transaction = connection.BeginTransaction();
            ExecuteNonQuery(connection, transaction, "DELETE FROM custom_input_methods;");

            foreach (var inputMethod in customInputMethods)
            {
                using var command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = """
                    INSERT INTO custom_input_methods (display_name, category, updated_at)
                    VALUES ($displayName, $category, CURRENT_TIMESTAMP);
                    """;
                command.Parameters.AddWithValue("$displayName", inputMethod.DisplayName);
                command.Parameters.AddWithValue("$category", inputMethod.Category);
                command.ExecuteNonQuery();
            }

            transaction.Commit();
        }
        catch
        {
            // Keep failures silent to avoid breaking the main workflow.
        }
    }

    private static void MigrateLegacyConfigJsonIfNeeded(SqliteConnection connection)
    {
        var values = ApplicationData.Current.LocalSettings.Values;
        var legacyConfigPath = Path.Combine(ApplicationData.Current.LocalFolder.Path, LegacyConfigFileName);
        if (!File.Exists(legacyConfigPath))
        {
            return;
        }

        var json = File.ReadAllText(legacyConfigPath, Encoding.UTF8);
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
        {
            return;
        }

        var selectedGamePath = ReadString(root, "SelectedGamePath");
        if (string.IsNullOrWhiteSpace(selectedGamePath))
        {
            selectedGamePath = ReadString(root, "GameDir");
        }

        if (!string.IsNullOrWhiteSpace(selectedGamePath))
        {
            values[SelectedGamePathKey] = selectedGamePath;
        }

        UpsertLegacyGamePaths(connection, root);
        UpsertLegacyInputMethods(connection, root);
        RenameMigratedLegacyConfig(legacyConfigPath);
    }

    private static void UpsertLegacyGamePaths(SqliteConnection connection, JsonElement root)
    {
        if (!root.TryGetProperty("GamePaths", out var gamePaths) || gamePaths.ValueKind != JsonValueKind.Array)
        {
            return;
        }

        using var transaction = connection.BeginTransaction();
        foreach (var gamePath in gamePaths.EnumerateArray())
        {
            if (gamePath.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var path = ReadString(gamePath, "Path")?.Trim();
            if (string.IsNullOrWhiteSpace(path))
            {
                continue;
            }

            var displayName = ReadString(gamePath, "Name");
            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = Path.GetFileName(path);
            }

            if (string.IsNullOrWhiteSpace(displayName))
            {
                displayName = path;
            }

            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = """
                INSERT INTO custom_game_paths (display_name, path, updated_at)
                VALUES ($displayName, $path, CURRENT_TIMESTAMP)
                ON CONFLICT(path) DO UPDATE SET
                    display_name = excluded.display_name,
                    updated_at = CURRENT_TIMESTAMP;
                """;
            command.Parameters.AddWithValue("$displayName", displayName);
            command.Parameters.AddWithValue("$path", path);
            command.ExecuteNonQuery();
        }

        transaction.Commit();
    }

    private static void UpsertLegacyInputMethods(SqliteConnection connection, JsonElement root)
    {
        if (!root.TryGetProperty("Ime", out var inputMethods) || inputMethods.ValueKind != JsonValueKind.Array)
        {
            return;
        }

        using var transaction = connection.BeginTransaction();
        foreach (var inputMethod in inputMethods.EnumerateArray())
        {
            if (inputMethod.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            var displayName = ReadString(inputMethod, "Name")?.Trim();
            if (string.IsNullOrWhiteSpace(displayName))
            {
                continue;
            }

            var category = NormalizeLegacyCategory(ReadString(inputMethod, "Category"));
            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = """
                INSERT INTO custom_input_methods (display_name, category, updated_at)
                VALUES ($displayName, $category, CURRENT_TIMESTAMP)
                ON CONFLICT(display_name) DO UPDATE SET
                    category = excluded.category,
                    updated_at = CURRENT_TIMESTAMP;
                """;
            command.Parameters.AddWithValue("$displayName", displayName);
            command.Parameters.AddWithValue("$category", category);
            command.ExecuteNonQuery();
        }

        transaction.Commit();
    }

    private static string NormalizeLegacyCategory(string? category) => category switch
    {
        "ChineseTraditional" => nameof(ImeCategory.ChineseTraditional),
        "Japanese" => nameof(ImeCategory.Japanese),
        _ => nameof(ImeCategory.ChineseSimplified)
    };

    private static string? ReadString(JsonElement element, string propertyName)
    {
        return element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString()
            : null;
    }

    private static void RenameMigratedLegacyConfig(string legacyConfigPath)
    {
        var migratedConfigPath = Path.Combine(ApplicationData.Current.LocalFolder.Path, MigratedLegacyConfigFileName);
        if (File.Exists(migratedConfigPath))
        {
            File.Delete(migratedConfigPath);
        }

        File.Move(legacyConfigPath, migratedConfigPath);
    }

    private static SqliteConnection OpenConnection()
    {
        var databasePath = Path.Combine(ApplicationData.Current.LocalFolder.Path, DatabaseFileName);
        var connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = databasePath
        }.ToString();

        var connection = new SqliteConnection(connectionString);
        connection.Open();
        return connection;
    }

    private static void ExecuteNonQuery(SqliteConnection connection, SqliteTransaction transaction, string commandText)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = commandText;
        command.ExecuteNonQuery();
    }

    private static void ExecuteNonQuery(SqliteConnection connection, string commandText)
    {
        using var command = connection.CreateCommand();
        command.CommandText = commandText;
        command.ExecuteNonQuery();
    }
}
