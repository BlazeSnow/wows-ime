using Microsoft.Data.Sqlite;
using System.Text;
using System.Text.Json;
using Windows.Storage;
using wows_ime.Core.Abstractions;
using wows_ime.Core.Models;
using wows_ime.Core.Rules;

namespace wows_ime.Windows;

public sealed class SettingsPersistence : ISettingsRepository
{
    private const int CurrentSchemaVersion = 1;
    private const string SchemaVersionKey = "Settings.SchemaVersion";
    private const string SelectedLanguageKey = "Settings.Language";
    private const string SelectedGamePathKey = "Game.SelectedPath";
    private const string LegacyConfigFileName = "config.json";
    private const string MigratedLegacyConfigFileName = "config.json.migrated";
    private const string DatabaseFileName = "settings.db";

    public void Initialize()
    {
        try
        {
            ApplicationData.Current.LocalSettings.Values[SchemaVersionKey] = CurrentSchemaVersion;
            using var connection = OpenConnection();
            ExecuteNonQuery(connection, """
                CREATE TABLE IF NOT EXISTS custom_game_paths (
                    display_name TEXT NOT NULL,
                    path TEXT PRIMARY KEY,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                """);
            ExecuteNonQuery(connection, """
                CREATE TABLE IF NOT EXISTS custom_input_methods (
                    display_name TEXT PRIMARY KEY,
                    category TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );
                """);
            EnsurePrimaryKeySchema(connection);
            MigrateLegacyConfigJsonIfNeeded(connection);
        }
        catch
        {
            // Local persistence failures must not prevent the application from starting.
        }
    }

    public string? LoadSelectedGamePath()
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

    public void SaveSelectedGamePath(string selectedGamePath)
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

    public string? LoadLanguageMode()
    {
        try
        {
            return ApplicationData.Current.LocalSettings.Values.TryGetValue(SelectedLanguageKey, out var value) &&
                   value is string language && LanguageRules.IsSupportedMode(language)
                ? language
                : null;
        }
        catch
        {
            return null;
        }
    }

    public void SaveLanguageMode(string languageMode)
    {
        try
        {
            ApplicationData.Current.LocalSettings.Values[SelectedLanguageKey] = LanguageRules.NormalizeMode(languageMode);
        }
        catch
        {
            // Keep failures silent to avoid breaking the settings workflow.
        }
    }

    public void ApplyLanguageMode()
    {
        try
        {
            var languageMode = LoadLanguageMode();
            if (languageMode is null)
            {
                var legacyLanguage = global::Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride?.Trim();
                languageMode = LanguageRules.IsSupportedMode(legacyLanguage) && legacyLanguage != LanguageRules.Automatic
                    ? legacyLanguage!
                    : LanguageRules.Automatic;
                SaveLanguageMode(languageMode);
            }

            global::Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride =
                languageMode == LanguageRules.Automatic ? string.Empty : languageMode;
        }
        catch
        {
            // Keep failures silent to avoid breaking application startup.
        }
    }

    public IReadOnlyList<PersistedGamePath> LoadCustomGamePaths()
    {
        var results = new List<PersistedGamePath>();
        try
        {
            using var connection = OpenConnection();
            using var command = connection.CreateCommand();
            command.CommandText = """
                SELECT display_name, path
                FROM custom_game_paths
                ORDER BY updated_at, path;
                """;
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                results.Add(new PersistedGamePath(reader.GetString(0), reader.GetString(1)));
            }
        }
        catch
        {
            // Return all successfully read values when a later read fails.
        }

        return results;
    }

    public void SaveCustomGamePaths(IEnumerable<PersistedGamePath> customGamePaths)
    {
        try
        {
            var paths = customGamePaths.ToList();
            using var connection = OpenConnection();
            using var transaction = connection.BeginTransaction();
            foreach (var path in paths)
            {
                using var command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = """
                    INSERT INTO custom_game_paths (display_name, path, updated_at)
                    VALUES ($displayName, $path, CURRENT_TIMESTAMP)
                    ON CONFLICT(path) DO UPDATE SET
                        display_name = excluded.display_name,
                        updated_at = CURRENT_TIMESTAMP;
                    """;
                command.Parameters.AddWithValue("$displayName", path.DisplayName);
                command.Parameters.AddWithValue("$path", path.Path);
                command.ExecuteNonQuery();
            }

            DeleteMissingRows(connection, transaction, "custom_game_paths", "path", paths.Select(item => item.Path).ToList());
            transaction.Commit();
        }
        catch
        {
            // Keep failures silent to avoid breaking the main workflow.
        }
    }

    public IReadOnlyList<PersistedInputMethod> LoadCustomInputMethods()
    {
        var results = new List<PersistedInputMethod>();
        try
        {
            using var connection = OpenConnection();
            using var command = connection.CreateCommand();
            command.CommandText = """
                SELECT display_name, category
                FROM custom_input_methods
                ORDER BY updated_at, display_name;
                """;
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                results.Add(new PersistedInputMethod(reader.GetString(0), reader.GetString(1)));
            }
        }
        catch
        {
            // Return all successfully read values when a later read fails.
        }

        return results;
    }

    public void SaveCustomInputMethods(IEnumerable<PersistedInputMethod> customInputMethods)
    {
        try
        {
            var inputMethods = customInputMethods.ToList();
            using var connection = OpenConnection();
            using var transaction = connection.BeginTransaction();
            foreach (var inputMethod in inputMethods)
            {
                using var command = connection.CreateCommand();
                command.Transaction = transaction;
                command.CommandText = """
                    INSERT INTO custom_input_methods (display_name, category, updated_at)
                    VALUES ($displayName, $category, CURRENT_TIMESTAMP)
                    ON CONFLICT(display_name) DO UPDATE SET
                        category = excluded.category,
                        updated_at = CURRENT_TIMESTAMP;
                    """;
                command.Parameters.AddWithValue("$displayName", inputMethod.DisplayName);
                command.Parameters.AddWithValue("$category", inputMethod.Category);
                command.ExecuteNonQuery();
            }

            DeleteMissingRows(connection, transaction, "custom_input_methods", "display_name", inputMethods.Select(item => item.DisplayName).ToList());
            transaction.Commit();
        }
        catch
        {
            // Keep failures silent to avoid breaking the main workflow.
        }
    }

    private static void DeleteMissingRows(SqliteConnection connection, SqliteTransaction transaction, string tableName, string keyColumnName, IReadOnlyList<string> valuesToKeep)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        if (valuesToKeep.Count == 0)
        {
            command.CommandText = $"DELETE FROM {tableName};";
        }
        else
        {
            var parameterNames = valuesToKeep.Select((_, index) => $"$value{index}").ToList();
            command.CommandText = $"DELETE FROM {tableName} WHERE {keyColumnName} NOT IN ({string.Join(", ", parameterNames)});";
            for (var i = 0; i < valuesToKeep.Count; i++)
            {
                command.Parameters.AddWithValue(parameterNames[i], valuesToKeep[i]);
            }
        }

        command.ExecuteNonQuery();
    }

    private static void MigrateLegacyConfigJsonIfNeeded(SqliteConnection connection)
    {
        var values = ApplicationData.Current.LocalSettings.Values;
        var legacyConfigPath = Path.Combine(ApplicationData.Current.LocalFolder.Path, LegacyConfigFileName);
        if (!File.Exists(legacyConfigPath))
        {
            return;
        }

        using var document = JsonDocument.Parse(File.ReadAllText(legacyConfigPath, Encoding.UTF8));
        var root = document.RootElement;
        if (root.ValueKind != JsonValueKind.Object)
        {
            return;
        }

        var selectedGamePath = ReadString(root, "SelectedGamePath") ?? ReadString(root, "GameDir");
        if (!string.IsNullOrWhiteSpace(selectedGamePath))
        {
            values[SelectedGamePathKey] = selectedGamePath;
        }

        UpsertLegacyGamePaths(connection, root);
        UpsertLegacyInputMethods(connection, root);
        RenameMigratedLegacyConfig(legacyConfigPath);
    }

    private static void EnsurePrimaryKeySchema(SqliteConnection connection)
    {
        if (!ColumnExists(connection, "custom_game_paths", "id") &&
            IsPrimaryKeyColumn(connection, "custom_game_paths", "path") &&
            !ColumnExists(connection, "custom_input_methods", "id") &&
            IsPrimaryKeyColumn(connection, "custom_input_methods", "display_name"))
        {
            return;
        }

        using var transaction = connection.BeginTransaction();
        MigrateCustomGamePathsTable(connection, transaction);
        MigrateCustomInputMethodsTable(connection, transaction);
        transaction.Commit();
    }

    private static void MigrateCustomGamePathsTable(SqliteConnection connection, SqliteTransaction transaction)
    {
        ExecuteNonQuery(connection, transaction, """
            CREATE TABLE IF NOT EXISTS custom_game_paths_new (
                display_name TEXT NOT NULL,
                path TEXT PRIMARY KEY,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            INSERT INTO custom_game_paths_new (display_name, path, created_at, updated_at)
            SELECT source.display_name, source.path, source.created_at, source.updated_at
            FROM custom_game_paths source
            WHERE source.path IS NOT NULL AND source.path <> ''
                AND source.rowid = (
                    SELECT latest.rowid FROM custom_game_paths latest
                    WHERE latest.path = source.path
                    ORDER BY latest.updated_at DESC, latest.rowid DESC LIMIT 1
                );
            DROP TABLE custom_game_paths;
            ALTER TABLE custom_game_paths_new RENAME TO custom_game_paths;
            """);
    }

    private static void MigrateCustomInputMethodsTable(SqliteConnection connection, SqliteTransaction transaction)
    {
        ExecuteNonQuery(connection, transaction, """
            CREATE TABLE IF NOT EXISTS custom_input_methods_new (
                display_name TEXT PRIMARY KEY,
                category TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            INSERT INTO custom_input_methods_new (display_name, category, created_at, updated_at)
            SELECT source.display_name, source.category, source.created_at, source.updated_at
            FROM custom_input_methods source
            WHERE source.display_name IS NOT NULL AND source.display_name <> ''
                AND source.rowid = (
                    SELECT latest.rowid FROM custom_input_methods latest
                    WHERE latest.display_name = source.display_name
                    ORDER BY latest.updated_at DESC, latest.rowid DESC LIMIT 1
                );
            DROP TABLE custom_input_methods;
            ALTER TABLE custom_input_methods_new RENAME TO custom_input_methods;
            """);
    }

    private static bool ColumnExists(SqliteConnection connection, string tableName, string columnName)
    {
        using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info({tableName});";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            if (string.Equals(reader.GetString(1), columnName, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static bool IsPrimaryKeyColumn(SqliteConnection connection, string tableName, string columnName)
    {
        using var command = connection.CreateCommand();
        command.CommandText = $"PRAGMA table_info({tableName});";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            if (string.Equals(reader.GetString(1), columnName, StringComparison.OrdinalIgnoreCase))
            {
                return reader.GetInt32(5) > 0;
            }
        }

        return false;
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
            displayName = string.IsNullOrWhiteSpace(displayName) ? Path.GetFileName(path) : displayName;
            displayName = string.IsNullOrWhiteSpace(displayName) ? path : displayName;
            UpsertGamePath(connection, transaction, displayName, path);
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
            var displayName = inputMethod.ValueKind == JsonValueKind.Object ? ReadString(inputMethod, "Name")?.Trim() : null;
            if (string.IsNullOrWhiteSpace(displayName))
            {
                continue;
            }

            using var command = connection.CreateCommand();
            command.Transaction = transaction;
            command.CommandText = """
                INSERT INTO custom_input_methods (display_name, category, updated_at)
                VALUES ($displayName, $category, CURRENT_TIMESTAMP)
                ON CONFLICT(display_name) DO UPDATE SET category = excluded.category, updated_at = CURRENT_TIMESTAMP;
                """;
            command.Parameters.AddWithValue("$displayName", displayName);
            command.Parameters.AddWithValue("$category", NormalizeLegacyCategory(ReadString(inputMethod, "Category")));
            command.ExecuteNonQuery();
        }

        transaction.Commit();
    }

    private static void UpsertGamePath(SqliteConnection connection, SqliteTransaction transaction, string displayName, string path)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            INSERT INTO custom_game_paths (display_name, path, updated_at)
            VALUES ($displayName, $path, CURRENT_TIMESTAMP)
            ON CONFLICT(path) DO UPDATE SET display_name = excluded.display_name, updated_at = CURRENT_TIMESTAMP;
            """;
        command.Parameters.AddWithValue("$displayName", displayName);
        command.Parameters.AddWithValue("$path", path);
        command.ExecuteNonQuery();
    }

    private static string NormalizeLegacyCategory(string? category) => category switch
    {
        "ChineseTraditional" => nameof(ImeCategory.ChineseTraditional),
        "Japanese" => nameof(ImeCategory.Japanese),
        _ => nameof(ImeCategory.ChineseSimplified)
    };

    private static string? ReadString(JsonElement element, string propertyName) =>
        element.TryGetProperty(propertyName, out var property) && property.ValueKind == JsonValueKind.String
            ? property.GetString()
            : null;

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
        var connection = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = databasePath }.ToString());
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
