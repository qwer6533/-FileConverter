#pragma once

#include "framework.h"
#include "resource.h"
#include <commdlg.h>
#include <shlobj.h>
#include <string>
#include <vector>
#include <thread>
#include <atomic>

// 文件类型
enum FileType {
    FILE_TYPE_DOCUMENT,
    FILE_TYPE_IMAGE,
    FILE_TYPE_AUDIO
};

// 转换器类
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

    // 添加工具路径获取方法
    std::wstring GetToolPath(const std::wstring& toolName);
};

// 全局变量
extern FileConverter g_Converter;
extern HWND g_hProgressBar;
extern HWND g_hStatusText;

// 函数声明
void CreateControls(HWND hWnd);
void UpdateFormatComboBoxes(HWND hWnd, FileType fileType);
void UpdateStatus(const std::wstring& status);
void BrowseInputFile(HWND hWnd);
void BrowseOutputFile(HWND hWnd);
void StartConversion(HWND hWnd);
std::wstring GetFileTypeName(FileType type);