#include "��Ϣ.h"

const WCHAR wchar_�ļ���׺[] = L".abc";
const WCHAR wchar_�ļ�����[] = L"MikuPy2001.abc";
const WCHAR wchar_�ļ�����[] = L"һ���Զ����ļ�(.recipe)�Ĵ������";

//��ProgID��ע��proplists�Է���ȡ��ע�ᣬ����С�����������ܴ�����ļ���չ����Ӧ�ó���ĳ�ͻ
//https://docs.microsoft.com/en-us/windows/win32/properties/building-property-handlers-property-lists

//���û���ͣ����Ŀ��ʱ�����Խ���ʾ����Ϣ���ϡ�
//��XPϵͳ��Ч
const WCHAR c_szRecipeInfoTip[] = L"prop:"
//"System.ItemType;"
//"System.Author;"
//"System.Rating;"
//"Microsoft.SampleRecipe.Difficulty;"
"System.Author"
;
//������ʾ�����ԶԻ������ϸ��Ϣѡ��ϡ������ļ�����֧�ֵ����Ե������б�
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
//������ʾ��Ԥ�������С�
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
//������ʾ����Ŀ����ͼ�Աߵ�Ԥ�� Pane�ı������򡣲����������Ϊ3�ˡ���������б�����Ķ��������������֣�������������Ŀ��
const WCHAR c_szRecipePreviewTitle[] = L"prop:"
//"System.Title;"
//"System.ItemType;"
"System.Author"
;