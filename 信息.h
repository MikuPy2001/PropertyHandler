#pragma once
#include <windows.h>



extern const WCHAR wchar_文件后缀[];
extern const WCHAR wchar_文件类型[];
extern const WCHAR wchar_文件描述[];


//在ProgID上注册proplists以方便取消注册，并最小化与其他可能处理该文件扩展名的应用程序的冲突
//https://docs.microsoft.com/en-us/windows/win32/properties/building-property-handlers-property-lists

//当用户悬停在项目上时，属性将显示在信息尖上。
//仅XP系统有效
extern const WCHAR c_szRecipeInfoTip[];
//属性显示在属性对话框的详细信息选项卡上。这是文件类型支持的属性的完整列表。
extern const WCHAR c_szRecipeFullDetails[];
//属性显示在预览窗格中。
extern const WCHAR c_szRecipePreviewDetails[];
//属性显示在项目缩略图旁边的预览 Pane的标题区域。参赛人数最多为3人。如果属性列表包含的多于最大允许的数字，则会忽略其余条目。
extern const WCHAR c_szRecipePreviewTitle[];