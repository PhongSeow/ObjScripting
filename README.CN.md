# ObjScripting
#### [English Readme](https://github.com/PhongSeow/ObjScripting/blob/master/README.md)
这是一个将 Microsoft Scripting Runtime 映射到 .net 平台的类库和 NuGet 包，可以在不同的 .net 框架下使用。
类库需要 %windir%\SysWOW64\scrrun.dll 支持，Windows 通常默认已安装好了。
使用这个类库，可以方便将 VB6 和 ASP 的代码升级到 .net 平台。
如果程序运行在 IIS，需要将应用程序池高级设置中的应用32位应用程序设置为 True 。

## ***目录和文件描述***

### Release

发布执行码目录，如果你不想看到源程序，你可以直接使用这个目录中的文件。

##### Release\DotNet\ObjScriptingLib
类库的DLL文件和 NuGet 包。

### Src

源码目录。

#### Src\DotNet\ObjScriptingLib

类库目录

#### Src\DotNet\ObjScriptingDemo

示例目录
