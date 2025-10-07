#pragma once

#include "framework.h"
#include "resource.h"
#include <commdlg.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <thread>
#include <atomic>

// �ļ�����
enum FileType {
    FILE_TYPE_DOCUMENT,
    FILE_TYPE_IMAGE,
    FILE_TYPE_AUDIO
};

// ת������
class FileConverter {
private:
    std::wstring inputFile;
    std::wstring outputFile;
    FileType currentFileType;
    std::wstring inputFormat;
    std::wstring outputFormat;
    std::atomic<bool> isConverting;
    std::wstring GetExeDirectory();

public:
    FileConverter();
    ~FileConverter();

    bool SetInputFile(const std::wstring& file);
    bool SetOutputFile(const std::wstring& file);
    void SetFileType(FileType type);
    void SetInputFormat(const std::wstring& format);
    void SetOutputFormat(const std::wstring& format);

    bool Convert();
    bool ConvertDocument();
    bool ConvertImage();
    bool ConvertAudio();

    bool IsConverting() const { return isConverting; }
    void StopConversion() { isConverting = false; }

    static std::vector<std::wstring> GetSupportedFormats(FileType type);
    static std::wstring GetFileExtension(const std::wstring& filename);
    static FileType DetectFileType(const std::wstring& filename);

    // ��ӹ���·����ȡ����
    std::wstring GetToolPath(const std::wstring& toolName);
};

// ȫ�ֱ���
extern FileConverter g_Converter;
extern HWND g_hProgressBar;
extern HWND g_hStatusText;

// ��������
void CreateControls(HWND hWnd);
void UpdateFormatComboBoxes(HWND hWnd, FileType fileType);
void UpdateStatus(const std::wstring& status);
void BrowseInputFile(HWND hWnd);
void BrowseOutputFile(HWND hWnd);
void StartConversion(HWND hWnd);
std::wstring GetFileTypeName(FileType type);