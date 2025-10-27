// FileConverter.cpp : 定义应用程序的入口点。
//

#include "framework.h"
#include "FileConverter.h"
#include <commdlg.h>
#include <shlobj.h>
#include <string>
#include <thread>
#include <Windows.h>
#define STB_IMAGE_IMPLEMENTATION
#define STB_IMAGE_WRITE_IMPLEMENTATION
#include "stb_image.h"
#include "stb_image_write.h"

#define MAX_LOADSTRING 100

// 全局变量:
HINSTANCE hInst;                                // 当前实例
WCHAR szTitle[MAX_LOADSTRING];                  // 标题栏文本
WCHAR szWindowClass[MAX_LOADSTRING];            // 主窗口类名

// 转换器实例
FileConverter g_Converter;
HWND g_hProgressBar = nullptr;
HWND g_hStatusText = nullptr;

// 此代码模块中包含的函数的前向声明:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
    _In_opt_ HINSTANCE hPrevInstance,
    _In_ LPWSTR    lpCmdLine,
    _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: 在此处放置代码。

    // 初始化全局字符串
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_FILECONVERTER, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // 执行应用程序初始化:
    if (!InitInstance(hInstance, nCmdShow))
    {
        return FALSE;
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_FILECONVERTER));

    MSG msg;

    // 主消息循环:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int)msg.wParam;
}

//
//  函数: MyRegisterClass()
//
//  目标: 注册窗口类。
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc = WndProc;
    wcex.cbClsExtra = 0;
    wcex.cbWndExtra = 0;
    wcex.hInstance = hInstance;
    wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_FILECONVERTER));
    wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wcex.lpszMenuName = MAKEINTRESOURCEW(IDC_FILECONVERTER);
    wcex.lpszClassName = szWindowClass;
    wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   函数: InitInstance(HINSTANCE, int)
//
//   目标: 保存实例句柄并创建主窗口
//
//   注释:
//
//        在此函数中，我们在全局变量中保存实例句柄并
//        创建和显示主程序窗口。
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
    hInst = hInstance; // 将实例句柄存储在全局变量中

    HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, 0, 800, 400, nullptr, nullptr, hInstance, nullptr);

    if (!hWnd)
    {
        return FALSE;
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    return TRUE;
}

// 创建控件
void CreateControls(HWND hWnd)
{
    // 文件类型选择
    CreateWindowW(L"Static", L"文件类型:", WS_VISIBLE | WS_CHILD,
        20, 20, 80, 20, hWnd, nullptr, hInst, nullptr);

    HWND hFileTypeCombo = CreateWindowW(L"ComboBox", L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST,
        100, 20, 200, 200, hWnd, (HMENU)IDC_FILE_TYPE_COMBO, hInst, nullptr);

    // 添加文件类型选项
    SendMessageW(hFileTypeCombo, CB_ADDSTRING, 0, (LPARAM)L"文档文件");
    SendMessageW(hFileTypeCombo, CB_ADDSTRING, 0, (LPARAM)L"图片文件");
    SendMessageW(hFileTypeCombo, CB_ADDSTRING, 0, (LPARAM)L"音频文件");
    SendMessageW(hFileTypeCombo, CB_SETCURSEL, 0, 0);

    // 输入文件
    CreateWindowW(L"Static", L"输入文件:", WS_VISIBLE | WS_CHILD,
        20, 60, 80, 20, hWnd, nullptr, hInst, nullptr);

    CreateWindowW(L"Edit", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
        100, 60, 400, 20, hWnd, (HMENU)IDC_INPUT_FILE_EDIT, hInst, nullptr);

    CreateWindowW(L"Button", L"浏览...", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        510, 60, 60, 20, hWnd, (HMENU)IDC_INPUT_FILE_BROWSE, hInst, nullptr);

    // 输出文件
    CreateWindowW(L"Static", L"输出文件:", WS_VISIBLE | WS_CHILD,
        20, 90, 80, 20, hWnd, nullptr, hInst, nullptr);

    CreateWindowW(L"Edit", L"", WS_VISIBLE | WS_CHILD | WS_BORDER | ES_AUTOHSCROLL,
        100, 90, 400, 20, hWnd, (HMENU)IDC_OUTPUT_FILE_EDIT, hInst, nullptr);

    CreateWindowW(L"Button", L"浏览...", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        510, 90, 60, 20, hWnd, (HMENU)IDC_OUTPUT_FILE_BROWSE, hInst, nullptr);

    // 输入格式
    CreateWindowW(L"Static", L"输入格式:", WS_VISIBLE | WS_CHILD,
        20, 130, 80, 20, hWnd, nullptr, hInst, nullptr);

    CreateWindowW(L"ComboBox", L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | WS_DISABLED,
        100, 130, 150, 200, hWnd, (HMENU)IDC_INPUT_FORMAT_COMBO, hInst, nullptr);

    // 输出格式
    CreateWindowW(L"Static", L"输出格式:", WS_VISIBLE | WS_CHILD,
        270, 130, 80, 20, hWnd, nullptr, hInst, nullptr);

    CreateWindowW(L"ComboBox", L"", WS_VISIBLE | WS_CHILD | CBS_DROPDOWNLIST | WS_DISABLED,
        350, 130, 150, 200, hWnd, (HMENU)IDC_OUTPUT_FORMAT_COMBO, hInst, nullptr);

    // 转换按钮
    CreateWindowW(L"Button", L"开始转换", WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        520, 130, 100, 25, hWnd, (HMENU)IDC_CONVERT_BUTTON, hInst, nullptr);

    // 进度条
    CreateWindowW(L"Static", L"进度:", WS_VISIBLE | WS_CHILD,
        20, 180, 80, 20, hWnd, nullptr, hInst, nullptr);

    g_hProgressBar = CreateWindowW(PROGRESS_CLASS, L"", WS_VISIBLE | WS_CHILD,
        100, 180, 400, 20, hWnd, (HMENU)IDC_PROGRESS_BAR, hInst, nullptr);

    SendMessage(g_hProgressBar, PBM_SETRANGE, 0, MAKELPARAM(0, 100));
    SendMessage(g_hProgressBar, PBM_SETPOS, 0, 0);

    // 状态文本
    g_hStatusText = CreateWindowW(L"Static", L"就绪", WS_VISIBLE | WS_CHILD,
        20, 220, 600, 20, hWnd, (HMENU)IDC_STATUS_STATIC, hInst, nullptr);
}

// 更新格式组合框
void UpdateFormatComboBoxes(HWND hWnd, FileType fileType)
{
    HWND hInputCombo = GetDlgItem(hWnd, IDC_INPUT_FORMAT_COMBO);
    HWND hOutputCombo = GetDlgItem(hWnd, IDC_OUTPUT_FORMAT_COMBO);

    // 清空组合框
    SendMessage(hInputCombo, CB_RESETCONTENT, 0, 0);
    SendMessage(hOutputCombo, CB_RESETCONTENT, 0, 0);

    // 获取支持的格式
    auto formats = FileConverter::GetSupportedFormats(fileType);

    for (const auto& format : formats) {
        SendMessage(hInputCombo, CB_ADDSTRING, 0, (LPARAM)format.c_str());
        SendMessage(hOutputCombo, CB_ADDSTRING, 0, (LPARAM)format.c_str());
    }

    if (formats.size() > 0) {
        SendMessage(hInputCombo, CB_SETCURSEL, 0, 0);
        SendMessage(hOutputCombo, CB_SETCURSEL, 0, 0);

        // 启用组合框
        EnableWindow(hInputCombo, TRUE);
        EnableWindow(hOutputCombo, TRUE);
    }
}

// 更新输出文件扩展名
void UpdateOutputFileExtension(HWND hWnd)
{
    wchar_t inputFile[260] = { 0 };
    wchar_t outputFile[260] = { 0 };

    GetDlgItemTextW(hWnd, IDC_INPUT_FILE_EDIT, inputFile, 260);
    GetDlgItemTextW(hWnd, IDC_OUTPUT_FILE_EDIT, outputFile, 260);

    // 如果输入文件为空，不进行处理
    if (wcslen(inputFile) == 0) {
        return;
    }

    // 获取输出格式
    wchar_t outputFormat[20] = { 0 };
    HWND hOutputCombo = GetDlgItem(hWnd, IDC_OUTPUT_FORMAT_COMBO);
    int outputIndex = SendMessage(hOutputCombo, CB_GETCURSEL, 0, 0);
    if (outputIndex != CB_ERR) {
        SendMessage(hOutputCombo, CB_GETLBTEXT, outputIndex, (LPARAM)outputFormat);
    }
    else {
        wcscpy_s(outputFormat, L"txt"); // 默认格式
    }

    std::wstring newOutputFile;

    // 如果输出文件为空或者与输入文件相关，则生成新的输出文件名
    if (wcslen(outputFile) == 0 ||
        std::wstring(outputFile).find(L"_converted.") != std::wstring::npos) {

        std::wstring inputFileStr = inputFile;
        size_t dotPos = inputFileStr.find_last_of(L'.');

        if (dotPos != std::wstring::npos) {
            newOutputFile = inputFileStr.substr(0, dotPos) + L"_converted." + outputFormat;
        }
        else {
            newOutputFile = inputFileStr + L"_converted." + outputFormat;
        }
    }
    else {
        // 保留用户自定义的文件名，只更改扩展名
        std::wstring outputFileStr = outputFile;
        size_t dotPos = outputFileStr.find_last_of(L'.');

        if (dotPos != std::wstring::npos) {
            newOutputFile = outputFileStr.substr(0, dotPos) + L"." + outputFormat;
        }
        else {
            newOutputFile = outputFileStr + L"." + outputFormat;
        }
    }

    SetDlgItemTextW(hWnd, IDC_OUTPUT_FILE_EDIT, newOutputFile.c_str());
}

// 更新状态
void UpdateStatus(const std::wstring& status)
{
    if (g_hStatusText) {
        SetWindowTextW(g_hStatusText, status.c_str());
    }
}

// 浏览输入文件
void BrowseInputFile(HWND hWnd)
{
    OPENFILENAMEW ofn;
    wchar_t szFile[260] = { 0 };

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hWnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = L"所有文件\0*.*\0文档文件\0*.docx;*.txt;*.rtf;*.html;*.md;*.epub\0图片文件\0*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff\0音频文件\0*.mp3;*.wav;*.flac;*.aac;*.ogg\0";
    ofn.nFilterIndex = 1;
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_FILEMUSTEXIST;

    if (GetOpenFileNameW(&ofn) == TRUE) {
        SetDlgItemTextW(hWnd, IDC_INPUT_FILE_EDIT, szFile);

        // 自动检测文件类型并更新格式
        FileType fileType = FileConverter::DetectFileType(szFile);
        HWND hFileTypeCombo = GetDlgItem(hWnd, IDC_FILE_TYPE_COMBO);
        SendMessage(hFileTypeCombo, CB_SETCURSEL, fileType, 0);

        UpdateFormatComboBoxes(hWnd, fileType);

        // 设置输入格式
        std::wstring ext = FileConverter::GetFileExtension(szFile);
        if (!ext.empty()) {
            HWND hInputCombo = GetDlgItem(hWnd, IDC_INPUT_FORMAT_COMBO);
            int count = SendMessage(hInputCombo, CB_GETCOUNT, 0, 0);
            for (int i = 0; i < count; i++) {
                wchar_t format[20];
                SendMessage(hInputCombo, CB_GETLBTEXT, i, (LPARAM)format);
                if (ext == format) {
                    SendMessage(hInputCombo, CB_SETCURSEL, i, 0);
                    break;
                }
            }
        }

        // 自动生成输出文件名
        UpdateOutputFileExtension(hWnd);
    }
}

// 浏览输出文件
void BrowseOutputFile(HWND hWnd)
{
    OPENFILENAMEW ofn;
    wchar_t szFile[260] = { 0 };

    // 获取当前输出文件路径作为默认值
    GetDlgItemTextW(hWnd, IDC_OUTPUT_FILE_EDIT, szFile, 260);

    ZeroMemory(&ofn, sizeof(ofn));
    ofn.lStructSize = sizeof(ofn);
    ofn.hwndOwner = hWnd;
    ofn.lpstrFile = szFile;
    ofn.nMaxFile = sizeof(szFile);
    ofn.lpstrFilter = L"所有文件\0*.*\0";
    ofn.Flags = OFN_PATHMUSTEXIST | OFN_OVERWRITEPROMPT;

    if (GetSaveFileNameW(&ofn) == TRUE) {
        SetDlgItemTextW(hWnd, IDC_OUTPUT_FILE_EDIT, szFile);
    }
}

// 开始转换
void StartConversion(HWND hWnd)
{
    if (g_Converter.IsConverting()) {
        MessageBoxW(hWnd, L"转换正在进行中，请等待完成。", L"提示", MB_OK | MB_ICONINFORMATION);
        return;
    }

    wchar_t inputFile[260], outputFile[260];
    GetDlgItemTextW(hWnd, IDC_INPUT_FILE_EDIT, inputFile, 260);
    GetDlgItemTextW(hWnd, IDC_OUTPUT_FILE_EDIT, outputFile, 260);

    if (wcslen(inputFile) == 0 || wcslen(outputFile) == 0) {
        MessageBoxW(hWnd, L"请选择输入文件和输出文件。", L"错误", MB_OK | MB_ICONERROR);
        return;
    }

    // 获取文件类型
    HWND hFileTypeCombo = GetDlgItem(hWnd, IDC_FILE_TYPE_COMBO);
    FileType fileType = (FileType)SendMessage(hFileTypeCombo, CB_GETCURSEL, 0, 0);

    // 获取输入输出格式
    wchar_t inputFormat[20], outputFormat[20];
    HWND hInputCombo = GetDlgItem(hWnd, IDC_INPUT_FORMAT_COMBO);
    HWND hOutputCombo = GetDlgItem(hWnd, IDC_OUTPUT_FORMAT_COMBO);

    int inputIndex = SendMessage(hInputCombo, CB_GETCURSEL, 0, 0);
    int outputIndex = SendMessage(hOutputCombo, CB_GETCURSEL, 0, 0);

    if (inputIndex == CB_ERR || outputIndex == CB_ERR) {
        MessageBoxW(hWnd, L"请选择输入格式和输出格式。", L"错误", MB_OK | MB_ICONERROR);
        return;
    }

    SendMessage(hInputCombo, CB_GETLBTEXT, inputIndex, (LPARAM)inputFormat);
    SendMessage(hOutputCombo, CB_GETLBTEXT, outputIndex, (LPARAM)outputFormat);

    // 设置转换器参数
    g_Converter.SetInputFile(inputFile);
    g_Converter.SetOutputFile(outputFile);
    g_Converter.SetFileType(fileType);
    g_Converter.SetInputFormat(inputFormat);
    g_Converter.SetOutputFormat(outputFormat);

    // 在新线程中执行转换
    std::thread conversionThread([hWnd]() {
        // 使用PostMessage进行线程安全的UI更新
        PostMessage(hWnd, WM_USER + 1, (WPARAM)L"转换中...", 0);
        EnableWindow(GetDlgItem(hWnd, IDC_CONVERT_BUTTON), FALSE);

        bool success = g_Converter.Convert();

        PostMessage(hWnd, WM_USER + 1, (WPARAM)(success ? L"转换完成！" : L"转换失败！"), 0);
        EnableWindow(GetDlgItem(hWnd, IDC_CONVERT_BUTTON), TRUE);

        if (success) {
            MessageBoxW(hWnd, L"文件转换成功完成！", L"成功", MB_OK | MB_ICONINFORMATION);
        }
        else {
            MessageBoxW(hWnd, L"文件转换失败，请检查文件格式和转换工具是否可用。", L"错误", MB_OK | MB_ICONERROR);
        }
        });

    conversionThread.detach();
}

//
//  函数: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  目标: 处理主窗口的消息。
//
//  WM_COMMAND  - 处理应用程序菜单
//  WM_PAINT    - 绘制主窗口
//  WM_DESTROY  - 发送退出消息并返回
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_CREATE:
        CreateControls(hWnd);
        break;

    case WM_COMMAND:
    {
        int wmId = LOWORD(wParam);
        // 分析菜单选择:
        switch (wmId)
        {
        case IDM_ABOUT:
            DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
            break;
        case IDM_EXIT:
            DestroyWindow(hWnd);
            break;
        case IDC_INPUT_FILE_BROWSE:
            BrowseInputFile(hWnd);
            break;
        case IDC_OUTPUT_FILE_BROWSE:
            BrowseOutputFile(hWnd);
            break;
        case IDC_CONVERT_BUTTON:
            StartConversion(hWnd);
            break;
        case IDC_FILE_TYPE_COMBO:
            if (HIWORD(wParam) == CBN_SELCHANGE) {
                HWND hCombo = (HWND)lParam;
                FileType fileType = (FileType)SendMessage(hCombo, CB_GETCURSEL, 0, 0);
                UpdateFormatComboBoxes(hWnd, fileType);
                UpdateOutputFileExtension(hWnd); // 文件类型改变时也更新输出文件后缀
            }
            break;
        case IDC_OUTPUT_FORMAT_COMBO:
            if (HIWORD(wParam) == CBN_SELCHANGE) {
                // 输出格式改变时更新输出文件后缀
                UpdateOutputFileExtension(hWnd);
            }
            break;
        default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
    }
    break;

    case WM_USER + 1: // 自定义消息用于线程安全的UI更新
        UpdateStatus((const wchar_t*)wParam);
        break;

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        // TODO: 在此处添加使用 hdc 的任何绘图代码...
        EndPaint(hWnd, &ps);
    }
    break;
    case WM_DESTROY:
        g_Converter.StopConversion();
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// "关于"框的消息处理程序。
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}

// 获取文件类型名称
std::wstring GetFileTypeName(FileType type)
{
    switch (type) {
    case FILE_TYPE_DOCUMENT:
        return L"文档文件";
    case FILE_TYPE_IMAGE:
        return L"图片文件";
    case FILE_TYPE_AUDIO:
        return L"音频文件";
    default:
        return L"未知文件";
    }
}

// FileConverter 类实现
FileConverter::FileConverter() : currentFileType(FILE_TYPE_DOCUMENT), isConverting(false)
{
}

FileConverter::~FileConverter()
{
    StopConversion();
}

std::wstring FileConverter::GetToolPath(const std::wstring& toolName)
{
    WCHAR exePath[MAX_PATH];
    GetModuleFileNameW(nullptr, exePath, MAX_PATH);
    std::wstring basePath = exePath;    //路径结构为打包路径而非调试路径
    size_t pos = basePath.find_last_of(L'\\');
    if (pos != std::wstring::npos) {
        basePath = basePath.substr(0, pos + 1);
    }
    return basePath + L"tools\\" + toolName;
}

bool FileConverter::SetInputFile(const std::wstring& file)
{
    inputFile = file;
    return true;
}

bool FileConverter::SetOutputFile(const std::wstring& file)
{
    outputFile = file;
    return true;
}

void FileConverter::SetFileType(FileType type)
{
    currentFileType = type;
}

void FileConverter::SetInputFormat(const std::wstring& format)
{
    inputFormat = format;
}

void FileConverter::SetOutputFormat(const std::wstring& format)
{
    outputFormat = format;
}

bool FileConverter::Convert()
{
    isConverting = true;
    bool result = false;

    switch (currentFileType) {
    case FILE_TYPE_DOCUMENT:
        result = ConvertDocument();
        break;
    case FILE_TYPE_IMAGE:
        result = ConvertImage();
        break;
    case FILE_TYPE_AUDIO:
        result = ConvertAudio();
        break;
    }

    isConverting = false;
    return result;
}

bool FileConverter::ConvertDocument()
{
    std::wstring pandocPath = GetToolPath(L"pandoc.exe");
    std::wstring iconvPath = GetToolPath(L"iconv.exe");

    // 检查工具是否存在
    if (_waccess(pandocPath.c_str(), 0) != 0) {
#ifdef _DEBUG
        OutputDebugStringW(L"Pandoc not found: ");
        OutputDebugStringW(pandocPath.c_str());
        OutputDebugStringW(L"\n");
#endif
        return false;
    }

    // 如果iconv不存在，直接使用pandoc（pandoc有内置编码处理）
    bool useIconv = (_waccess(iconvPath.c_str(), 0) == 0);

#ifdef _DEBUG
    OutputDebugStringW(L"Use iconv: ");
    OutputDebugStringW(useIconv ? L"true" : L"false");
    OutputDebugStringW(L"\n");
#endif

    // 映射输入输出格式
    std::wstring pandocInputFormat = inputFormat;
    std::wstring pandocOutputFormat = outputFormat;

    // 输入格式映射
    if (inputFormat == L"txt") {
        pandocInputFormat = L"markdown";
    }
    else if (inputFormat == L"md") {
        pandocInputFormat = L"markdown";
    }

    // 输出格式映射
    if (outputFormat == L"txt") {
        pandocOutputFormat = L"plain";
    }
    else if (outputFormat == L"md") {
        pandocOutputFormat = L"markdown";
    }

    std::wstring finalInputFile = inputFile;
    std::wstring tempFile;

    // 如果iconv可用，尝试检测编码并转换
    if (useIconv) {
        tempFile = outputFile + L".temp_utf8";

        // 常见编码列表
        std::vector<std::wstring> encodings = {
            // Unicode系列 
            L"UTF-8", L"UTF-16", L"UTF-16LE", L"UTF-16BE",

            // 中文编码
            L"GBK", L"GB2312", L"GB18030", L"BIG5",
            L"CP936", L"CP950",

            // 日语编码
            L"SHIFT_JIS", L"CP932", L"EUC-JP", L"ISO-2022-JP",

            // 韩语编码
            L"CP949", L"EUC-KR",

            // 其他常见编码
            L"ISO-8859-1", L"WINDOWS-1252"
        };

        bool conversionSuccess = false;

        for (const auto& encoding : encodings) {
            // 修复iconv命令行格式
            std::wstring iconvCommand = L"\"" + iconvPath + L"\" -f " + encoding +
                L" -t UTF-8 \"" + inputFile +
                L"\" -o \"" + tempFile + L"\"";

            STARTUPINFOW si = { sizeof(si) };
            PROCESS_INFORMATION pi;
            wchar_t cmdLine[1024];
            wcscpy_s(cmdLine, iconvCommand.c_str());

#ifdef _DEBUG
            OutputDebugStringW(L"尝试编码: ");
            OutputDebugStringW(encoding.c_str());
            OutputDebugStringW(L" - 命令: ");
            OutputDebugStringW(cmdLine);
            OutputDebugStringW(L"\n");
#endif

            BOOL processCreated = CreateProcessW(
                nullptr,
                cmdLine,
                nullptr,
                nullptr,
                FALSE,
                CREATE_NO_WINDOW,
                nullptr,
                nullptr,
                &si,
                &pi
            );

            if (!processCreated) {
#ifdef _DEBUG
                OutputDebugStringW(L"创建进程失败\n");
#endif
                continue;
            }

            WaitForSingleObject(pi.hProcess, 10000); // 10秒超时
            DWORD exitCode;
            GetExitCodeProcess(pi.hProcess, &exitCode);
            CloseHandle(pi.hProcess);
            CloseHandle(pi.hThread);

            // 检查转换是否成功
            if (exitCode == 0 && _waccess(tempFile.c_str(), 0) == 0) {
                FILE* testFile = nullptr;
                errno_t openResult = _wfopen_s(&testFile, tempFile.c_str(), L"rb");

                if (openResult == 0 && testFile != nullptr) {
                    fseek(testFile, 0, SEEK_END);
                    long fileSize = ftell(testFile);
                    fclose(testFile);

                    if (fileSize > 0) {
                        conversionSuccess = true;
                        finalInputFile = tempFile;

#ifdef _DEBUG
                        OutputDebugStringW(L"成功使用编码: ");
                        OutputDebugStringW(encoding.c_str());
                        OutputDebugStringW(L"\n");
#endif
                        break;
                    }
                }
            }

            // 删除失败的临时文件
            DeleteFileW(tempFile.c_str());
        }

        // 如果所有编码尝试都失败，回退到直接使用原始文件
        if (!conversionSuccess) {
#ifdef _DEBUG
            OutputDebugStringW(L"所有编码尝试失败，使用原始文件\n");
#endif
            finalInputFile = inputFile;
        }
    }

    // 构建pandoc命令
    std::wstring pandocCommand = L"\"" + pandocPath + L"\" \"" + finalInputFile +
        L"\" -f " + pandocInputFormat + L" -t " + pandocOutputFormat + L" -o \"" + outputFile + L"\"";

#ifdef _DEBUG
    OutputDebugStringW(L"执行pandoc命令: ");
    OutputDebugStringW(pandocCommand.c_str());
    OutputDebugStringW(L"\n");
#endif

    // 执行pandoc
    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi;
    wchar_t cmdLine[1024];
    wcscpy_s(cmdLine, pandocCommand.c_str());

    BOOL success = CreateProcessW(
        nullptr,
        cmdLine,
        nullptr,
        nullptr,
        FALSE,
        CREATE_NO_WINDOW,
        nullptr,
        nullptr,
        &si,
        &pi
    );

    bool result = false;

    if (success) {
        WaitForSingleObject(pi.hProcess, 30000); // 30秒超时
        DWORD exitCode;
        GetExitCodeProcess(pi.hProcess, &exitCode);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        result = (exitCode == 0);
    }

    // 清理临时文件
    if (!tempFile.empty() && _waccess(tempFile.c_str(), 0) == 0) {
        DeleteFileW(tempFile.c_str());
    }

    return result;
}

bool FileConverter::ConvertImage()
{
    // 使用stb_image进行图片转换
    int width, height, channels;

    // 将宽字符串转换为多字节字符串（UTF-8）
    int inputSize = WideCharToMultiByte(CP_UTF8, 0, inputFile.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::vector<char> inputBuffer(inputSize);
    WideCharToMultiByte(CP_UTF8, 0, inputFile.c_str(), -1, inputBuffer.data(), inputSize, nullptr, nullptr);

    int outputSize = WideCharToMultiByte(CP_UTF8, 0, outputFile.c_str(), -1, nullptr, 0, nullptr, nullptr);
    std::vector<char> outputBuffer(outputSize);
    WideCharToMultiByte(CP_UTF8, 0, outputFile.c_str(), -1, outputBuffer.data(), outputSize, nullptr, nullptr);

    // 加载图像
    unsigned char* image_data = stbi_load(inputBuffer.data(), &width, &height, &channels, 0);
    if (!image_data) {
        return false;
    }

    bool success = false;

    // 根据输出格式选择保存方式
    if (outputFormat == L"jpg" || outputFormat == L"jpeg") {
        success = stbi_write_jpg(outputBuffer.data(), width, height, channels, image_data, 90);
    }
    else if (outputFormat == L"png") {
        success = stbi_write_png(outputBuffer.data(), width, height, channels, image_data, width * channels);
    }
    else if (outputFormat == L"bmp") {
        success = stbi_write_bmp(outputBuffer.data(), width, height, channels, image_data);
    }
    else if (outputFormat == L"tga") {
        success = stbi_write_tga(outputBuffer.data(), width, height, channels, image_data);
    }
    else {
        // 不支持的格式，默认保存为PNG
        success = stbi_write_png(outputBuffer.data(), width, height, channels, image_data, width * channels);
    }

    // 释放图像数据
    stbi_image_free(image_data);

    return success;
}

bool FileConverter::ConvertAudio()
{
    std::wstring toolPath = GetToolPath(L"ffmpeg.exe");

    // 检查工具是否存在
    if (_waccess(toolPath.c_str(), 0) != 0) {
        return false;
    }

    std::wstring command = L"\"" + toolPath + L"\" -i \"" + inputFile +
        L"\" \"" + outputFile + L"\"";

    // 执行命令
    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi;

    wchar_t cmdLine[1024];
    wcscpy_s(cmdLine, command.c_str());

    BOOL success = CreateProcessW(
        nullptr,
        cmdLine,
        nullptr,
        nullptr,
        FALSE,
        0,
        nullptr,
        nullptr,
        &si,
        &pi
    );

    if (success) {
        WaitForSingleObject(pi.hProcess, INFINITE);
        DWORD exitCode;
        GetExitCodeProcess(pi.hProcess, &exitCode);
        CloseHandle(pi.hProcess);
        CloseHandle(pi.hThread);
        return exitCode == 0;
    }

    return false;
}

std::vector<std::wstring> FileConverter::GetSupportedFormats(FileType type)
{
    std::vector<std::wstring> formats;

    switch (type) {
    case FILE_TYPE_DOCUMENT:
        formats = { L"docx", L"txt", L"rtf", L"html", L"md",L"epub"};
        break;
    case FILE_TYPE_IMAGE:
        formats = { L"jpg", L"jpeg", L"png", L"bmp", L"gif", L"tiff" };
        break;
    case FILE_TYPE_AUDIO:
        formats = { L"mp3", L"wav", L"flac", L"aac", L"ogg" };
        break;
    }

    return formats;
}

std::wstring FileConverter::GetFileExtension(const std::wstring& filename)
{
    size_t dotPos = filename.find_last_of(L'.');
    if (dotPos != std::wstring::npos) {
        std::wstring ext = filename.substr(dotPos + 1);
        // 转换为小写以进行不区分大小写的比较
        for (auto& c : ext) {
            c = towlower(c);
        }
        return ext;
    }
    return L"";
}

FileType FileConverter::DetectFileType(const std::wstring& filename)
{
    std::wstring ext = GetFileExtension(filename);

    // 文档格式
    if (ext == L"docx" || ext == L"txt" || ext == L"rtf" || 
        ext == L"html" || ext == L"md" || ext == L"epub") {
        return FILE_TYPE_DOCUMENT;
    }
    // 图片格式
    else if (ext == L"jpg" || ext == L"jpeg" || ext == L"png" ||
        ext == L"bmp" || ext == L"gif" || ext == L"tiff") {
        return FILE_TYPE_IMAGE;
    }
    // 音频格式
    else if (ext == L"mp3" || ext == L"wav" || ext == L"flac" ||
        ext == L"aac" || ext == L"ogg") {
        return FILE_TYPE_AUDIO;
    }

    return FILE_TYPE_DOCUMENT; // 默认
}
