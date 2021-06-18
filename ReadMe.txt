这是一个修改自官方的代码
将编译为一个DLL
你需要同时编译32位以及64位
然后
regsvr32.exe 你的32.dll
regsvr32.exe 你的64.dll
最后新建一个 .abc文件
输入以下内容
<?xml version="1.0" encoding="UTF-8"?>
<Temp>
    <Title>中文 Title</Title>
    <Author>Author 1</Author>
    <Author>Author 2</Author>
</Temp>
保存为UTF-8格式
然后将文件夹视图切换为 详细信息
右键标题栏,勾选 标题 作者
即可看到效果

建议使用 
regsvr32.exe /u 你的32.dll
regsvr32.exe /u 你的64.dll
卸载你的旧DLL
然后再更新新DLL


Recipe Property Handler Sample
================================
Demonstrates the implementation of property handler for custom file format for recipes (.recipe); this format supports several properties natively, and also supports reading and writing arbitrary properties and custom schema.  The sample also includes a custom property schema which defines properties that are germane to the .recipe file format (e.g. not represented by any properties in the System property schema).


Sample Language Implementations
===============================
C++

Files:
=============================================
Dll.cpp
RegisterExtension.cpp
RegisterExtension.h
RecipePropertyHandler.cpp
RecipePropertyHandler.def
RecipePropertyHandler.rc
RecipePropertyHandler.sln
RecipePropertyHandler.vcproj
RecipeProperies.propdesc
PerfectSteaks.recipe


To build the sample using the command prompt:
=============================================
     1. Open the Command Prompt window and navigate to the RecipePropertyHandler directory.
     2. Type msbuild RecipePropertyHandler.sln.


To build the sample using Visual Studio (preferred method):
===========================================================
     1. Open Windows Explorer and navigate to the RecipePropertyHandler directory.
     2. Double-click the icon for the RecipePropertyHandler.sln file to open the file in Visual Studio.
     3. In the Build menu, select Build Solution. The application will be built in the default Win32\Debug or Win32\Release directory (or x64\...).


To run the sample:
=================
     1. Navigate to the directory that contains the new dll using the command prompt.
     2. Type 'regsvr32.exe RecipePropertyHandler.dll'
     3. (Optional) Type 'prop.exe schema register RecipeProperties.propdesc' to register the custom .recipe properties in the property schema.
        NOTE: prop.exe can be obtained from http://www.codeplex.com/prop.  The PropertySchema SDK sample application can also be used to register the schema.

To remove the sample:
=================
     1. Navigate to the directory that contains the dll using the command prompt.
     2. Type regsvr32.exe /u RecipePropertyHandler.dll
     3. Type 'prop.exe schema unregister RecipeProperties.propdesc'
