# AI编程指导文件

1. 默认禁止对本文件进行修改；只有用户明确授权时才可以修订
2. 所有文件都以UTF-8格式存储

## 程序介绍

本程序针对游戏《战舰世界》的中文输入法的配置文件，目的为修订配置文件以支持更多的中文输入法。

## 配置文件示例

示例的配置文件见仓库根目录的ime_config.xml文件

## 程序逻辑

1. 扫描并列出用户的输入法
2. 由用户勾选要添加的输入法，并选择各个选定的输入法类型（中文简体、繁体或日文）
3. 支持添加自定义游戏路径和自定义输入法
4. 删除自定义游戏路径或自定义输入法前，需要弹窗确认
5. 检查配置文件目录是否已有ime_config.xml：若有，则提示用户覆盖；若无，由用户确认添加

## 配置文件目录

1. 需要由用户选定游戏根目录
2. 如果为Steam下载的游戏，默认是：`C:\Program Files (x86)\Steam\steamapps\common\World of Warships`
3. 如果是莱斯塔启动器下载的游戏，默认是：`C:\Games\Korabli`
4. 如果是360启动器下载的游戏，默认是：`C:\Games\World_of_Warships_CN360`
5. 也存在大量不是Steam下载的用户，所以需要提供选择文件夹功能
6. 用户选择游戏根目录后，需要确认游戏根目录存在游戏exe程序
7. 游戏exe程序一般是`WorldOfWarships.exe`，对于俄罗斯服玩家，exe程序是`Korabli.exe`
8. 确认游戏根目录后，本程序的输入法配置文件目录为`\bin\8842736\res_mods\ime_config.xml`，其中的`8842736`是游戏多版本中的一个版本；程序当前会为`bin`目录下所有数字版本目录写入输入法配置文件
9. 注意，`8842736`只是游戏版本示例，不允许以此为回退方案

## 持久化配置

1. 软件最终发布至Microsoft Store，本地测试也使用VS打包运行，因此按打包应用处理
2. 简单状态使用`ApplicationData.Current.LocalSettings`
   1. `Settings.SchemaVersion`：当前配置结构版本
   2. `Game.SelectedPath`：当前选中的游戏根目录
3. 列表数据使用SQLite，数据库文件位于`ApplicationData.Current.LocalFolder\settings.db`
4. SQLite表结构如下：
   1. `custom_game_paths`：自定义游戏路径，`path`为主键，字段包括`display_name`、`path`、`created_at`、`updated_at`
   2. `custom_input_methods`：自定义输入法，`display_name`为主键，字段包括`display_name`、`category`、`created_at`、`updated_at`
5. 旧版本曾使用`ApplicationData.Current.LocalFolder\config.json`；启动时如果发现该文件，需要迁移到新存储方式
6. 旧`config.json`迁移成功后，应重命名为`config.json.migrated`，不使用LocalSettings记录迁移标记
7. 迁移旧数据时：
   1. `SelectedGamePath`或`GameDir`迁移到`Game.SelectedPath`
   2. `GamePaths`迁移到`custom_game_paths`
   3. `Ime`迁移到`custom_input_methods`

## 程序架构

1. 程序使用winui 3架构
2. 最终打包发布至Microsoft Store
3. 项目使用单项目MSIX打包，并启用剪裁；`ILLink.Descriptors.xml`用于保留`WinRT.Runtime`
4. 主要代码结构：
   1. `Views\MainPage.xaml`和`Views\MainPage.xaml.cs`：主界面和页面交互逻辑
   2. `Views\InputMethodItem.cs`：输入法列表绑定模型
   3. `Views\GamePathOption.cs`：游戏路径列表绑定模型
   4. `Views\ImeCategory.cs`：输入法分类枚举
   5. `Services\InputMethodScanner.cs`：TSF/COM输入法扫描逻辑
   6. `Services\SettingsPersistence.cs`：LocalSettings、SQLite和旧配置迁移逻辑
   7. `Services\GameConfigService.cs`：游戏目录校验、目标配置路径解析、ime_config.xml写入逻辑
   8. `Services\AppResources.cs`：资源字符串读取和格式化
