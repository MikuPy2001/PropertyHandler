#include "信息.h"

const WCHAR wchar_文件后缀[] = L".abc";
const WCHAR wchar_文件类型[] = L"MikuPy2001.abc";
const WCHAR wchar_文件描述[] = L"一种自定义文件(.recipe)的处理程序";

//在ProgID上注册proplists以方便取消注册，并最小化与其他可能处理该文件扩展名的应用程序的冲突
//https://docs.microsoft.com/en-us/windows/win32/properties/building-property-handlers-property-lists

//当用户悬停在项目上时，属性将显示在信息尖上。
//仅XP系统有效
const WCHAR c_szRecipeInfoTip[] = L"prop:"
//"System.ItemType;"
//"System.Author;"
//"System.Rating;"
//"Microsoft.SampleRecipe.Difficulty;"
"System.Author"
;
//属性显示在属性对话框的详细信息选项卡上。这是文件类型支持的属性的完整列表。
const WCHAR c_szRecipeFullDetails[] = L"prop:"
"System.Title;"
//"System.PropGroup.Description;"
//"System.Title;"
//"System.Author;"
//"System.Comment;"
//"System.Keywords;"
//"System.Rating;"
//"Microsoft.SampleRecipe.Difficulty;"
//"System.PropGroup.FileSystem;"
//"System.ItemNameDisplay;"
//"System.ItemType;"
//"System.ItemFolderPathDisplay;"
//"System.Size;"
//"System.DateCreated;"
//"System.DateModified;"
//"System.DateAccessed;"
//"System.FileAttributes;"
//"System.OfflineAvailability;"
//"System.OfflineStatus;"
//"System.SharedWith;"
//"System.FileOwner;"
//"System.ComputerName;"
"System.Author"
;
//属性显示在预览窗格中。
const WCHAR c_szRecipePreviewDetails[] = L"prop:"
//"System.DateChanged;"
//"System.Author;"
//"System.Keywords;"
//"Microsoft.SampleRecipe.Difficulty;"
//"System.Rating;"
//"System.Comment;"
//"System.Size;"
//"System.ItemFolderPathDisplay;"
//"System.DateCreated;"
"System.Author"
;
//属性显示在项目缩略图旁边的预览 Pane的标题区域。参赛人数最多为3人。如果属性列表包含的多于最大允许的数字，则会忽略其余条目。
const WCHAR c_szRecipePreviewTitle[] = L"prop:"
//"System.Title;"
//"System.ItemType;"
"System.Author"
;