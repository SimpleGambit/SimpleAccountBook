using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using SimpleAccountBook.Models;

namespace SimpleAccountBook.Services;

public class AppStateService
{
    private readonly string _stateDirectory;
    private readonly string _stateFilePath;
    private readonly JsonSerializerOptions _serializerOptions;
    private readonly SemaphoreSlim _gate = new(1, 1);

    public AppStateService()
    {
        _stateDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SimpleAccountBook");
        Directory.CreateDirectory(_stateDirectory);
        _stateFilePath = Path.Combine(_stateDirectory, "app_state.json");
        _serializerOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            PropertyNameCaseInsensitive = true
        };
    }

    public async Task SaveAsync(AppState state)
    {
        if (state is null)
        {
            throw new ArgumentNullException(nameof(state));
        }

        await _gate.WaitAsync().ConfigureAwait(false);
        try
        {
            var json = JsonSerializer.Serialize(state, _serializerOptions);
            await File.WriteAllTextAsync(_stateFilePath, json).ConfigureAwait(false);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Debug.WriteLine($"[AppStateService] 상태 저장 실패: {ex}");
            throw;
        }
        finally
        {
            _gate.Release();
        }
    }

    public async Task<AppState?> LoadAsync()
    {
        await _gate.WaitAsync().ConfigureAwait(false);
        try
        {
            if (!File.Exists(_stateFilePath))
            {
                return null;
            }

            var json = await File.ReadAllTextAsync(_stateFilePath).ConfigureAwait(false);
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }

            return JsonSerializer.Deserialize<AppState>(json, _serializerOptions);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or JsonException)
        {
            Debug.WriteLine($"[AppStateService] 상태 불러오기 실패: {ex}");
            return null;
        }
        finally
        {
            _gate.Release();
        }
    }

    public bool HasSavedState => File.Exists(_stateFilePath);
}