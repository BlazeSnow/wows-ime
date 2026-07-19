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
    private const string SelectedLanguageKey = "Settings.Language";
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
            // Keep the app usable even when the local database cannot be initialized.
        }
    }

    internal static string? LoadSelectedLanguage()
    {
        try
        {
            return ApplicationData.Current.LocalSettings.Values.TryGetValue(SelectedLanguageKey, out var value) &&
                   value is string language &&
                   IsSupportedLanguage(language)
                ? language
                : null;
        }
        catch
        {
            return null;
        }
    }

    internal static void SaveSelectedLanguage(string language)
    {
        try
        {
            ApplicationData.Current.LocalSettings.Values[SelectedLanguageKey] = IsSupportedLanguage(language) ? language : "auto";
        }
        catch
        {
            // Keep failures silent to avoid breaking the settings workflow.
        }
    }

    internal static void ApplySelectedLanguage()
    {
        try
        {
            var language = LoadSelectedLanguage();
            if (language is null)
            {
                var legacyLanguage = Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride?.Trim();
                string migratedLanguage = IsSupportedLanguage(legacyLanguage) && legacyLanguage != "auto" ? legacyLanguage! : "auto";
                SaveSelectedLanguage(migratedLanguage);
                language = migratedLanguage;
            }

            Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = language == "auto" ? string.Empty : language;
        }
        catch
        {
            // Keep the app usable even when the language preference cannot be applied.
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
            return results;
        }

        return results;
    }

    internal static void SaveCustomGamePaths(IEnumerable<PersistedGamePath> customGamePaths)
    {
        try
        {
            var pathItems = customGamePaths.ToList();
            using var connection = OpenConnection();
            using var transaction = connection.BeginTransaction();

            foreach (var path in pathItems)
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

            DeleteMissingGamePaths(connection, transaction, pathItems.Select(item => item.Path).ToList());
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
            return results;
        }

        return results;
    }

    internal static void SaveCustomInputMethods(IEnumerable<PersistedInputMethod> customInputMethods)
    {
        try
        {
            var inputMethodItems = customInputMethods.ToList();
            using var connection = OpenConnection();
            using var transaction = connection.BeginTransaction();

            foreach (var inputMethod in inputMethodItems)
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

            DeleteMissingInputMethods(connection, transaction, inputMethodItems.Select(item => item.DisplayName).ToList());
            transaction.Commit();
        }
        catch
        {
            // Keep failures silent to avoid breaking the main workflow.
        }
    }

    private static void DeleteMissingGamePaths(SqliteConnection connection, SqliteTransaction transaction, IReadOnlyList<string> pathsToKeep)
    {
        DeleteMissingRows(connection, transaction, "custom_game_paths", "path", pathsToKeep);
    }

    private static void DeleteMissingInputMethods(SqliteConnection connection, SqliteTransaction transaction, IReadOnlyList<string> displayNamesToKeep)
    {
        DeleteMissingRows(connection, transaction, "custom_input_methods", "display_name", displayNamesToKeep);
    }

    private static void DeleteMissingRows(
        SqliteConnection connection,
        SqliteTransaction transaction,
        string tableName,
        string keyColumnName,
        IReadOnlyList<string> valuesToKeep)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        if (valuesToKeep.Count == 0)
        {
            command.CommandText = $"DELETE FROM {tableName};";
            command.ExecuteNonQuery();
            return;
        }

        var parameterNames = valuesToKeep
            .Select((_, index) => $"$value{index}")
            .ToList();
        command.CommandText = $"DELETE FROM {tableName} WHERE {keyColumnName} NOT IN ({string.Join(", ", parameterNames)});";
        for (var i = 0; i < valuesToKeep.Count; i++)
        {
            command.Parameters.AddWithValue(parameterNames[i], valuesToKeep[i]);
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
            """);

        ExecuteNonQuery(connection, transaction, """
            INSERT INTO custom_game_paths_new (display_name, path, created_at, updated_at)
            SELECT source.display_name, source.path, source.created_at, source.updated_at
            FROM custom_game_paths source
            WHERE source.path IS NOT NULL
                AND source.path <> ''
                AND source.rowid = (
                    SELECT latest.rowid
                    FROM custom_game_paths latest
                    WHERE latest.path = source.path
                    ORDER BY latest.updated_at DESC, latest.rowid DESC
                    LIMIT 1
                );
            """);

        ExecuteNonQuery(connection, transaction, "DROP TABLE custom_game_paths;");
        ExecuteNonQuery(connection, transaction, "ALTER TABLE custom_game_paths_new RENAME TO custom_game_paths;");
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
            """);

        ExecuteNonQuery(connection, transaction, """
            INSERT INTO custom_input_methods_new (display_name, category, created_at, updated_at)
            SELECT source.display_name, source.category, source.created_at, source.updated_at
            FROM custom_input_methods source
            WHERE source.display_name IS NOT NULL
                AND source.display_name <> ''
                AND source.rowid = (
                    SELECT latest.rowid
                    FROM custom_input_methods latest
                    WHERE latest.display_name = source.display_name
                    ORDER BY latest.updated_at DESC, latest.rowid DESC
                    LIMIT 1
                );
            """);

        ExecuteNonQuery(connection, transaction, "DROP TABLE custom_input_methods;");
        ExecuteNonQuery(connection, transaction, "ALTER TABLE custom_input_methods_new RENAME TO custom_input_methods;");
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

    internal static string ResolveLanguage(string? language)
    {
        if (language is "zh-Hans" or "zh-Hant" or "ja")
        {
            return language;
        }

        try
        {
            foreach (var preferredLanguage in Windows.System.UserProfile.GlobalizationPreferences.Languages)
            {
                if (IsSimplifiedChinese(preferredLanguage))
                {
                    return "zh-Hans";
                }

                if (IsTraditionalChinese(preferredLanguage))
                {
                    return "zh-Hant";
                }

                if (preferredLanguage.Equals("ja", StringComparison.OrdinalIgnoreCase) ||
                    preferredLanguage.StartsWith("ja-", StringComparison.OrdinalIgnoreCase))
                {
                    return "ja";
                }
            }
        }
        catch
        {
            // Use the packaged default resource language when system preferences are unavailable.
        }

        return "zh-Hans";
    }

    private static bool IsSimplifiedChinese(string language) =>
        language.Equals("zh-Hans", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-Hans-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-CN", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-CN-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-SG", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-SG-", StringComparison.OrdinalIgnoreCase);

    private static bool IsTraditionalChinese(string language) =>
        language.Equals("zh-Hant", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-Hant-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-TW", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-TW-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-HK", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-HK-", StringComparison.OrdinalIgnoreCase) ||
        language.Equals("zh-MO", StringComparison.OrdinalIgnoreCase) ||
        language.StartsWith("zh-MO-", StringComparison.OrdinalIgnoreCase);

    private static bool IsSupportedLanguage(string? language) => language is "auto" or "zh-Hans" or "zh-Hant" or "ja";

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
