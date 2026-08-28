# AI 编程指导

1. 本文件是 AI 编程的基准指导，默认禁止修改；只有用户明确授权时才可以修订
2. 所有文件都以 UTF-8 格式存储
3. 实际开发中需要注意终端的GBK和UTF-8的问题
4. 业务逻辑、持久化、架构、构建发布、本地化等详细内容见 [DEVELOPMENT.md](DEVELOPMENT.md)；DEVELOPMENT.md 鼓励随开发进度同步修改，代码或流程变更时无需额外授权即可更新

## 项目简介

本程序是《战舰世界》中文输入法配置文件修改器（WinUI 3 桌面应用），通过向游戏目录写入 `ime_config.xml` 使游戏支持更多中文输入法，最终发布至 Microsoft Store。

## 关键约定

- 解决方案 `wows-ime.slnx` 为三个项目：根主项目（`App.xaml.cs`、`PageHost.cs`，应用入口与组合根）+ `wows-ime.Core`（无 UI 核心逻辑：输入法扫描、持久化、游戏配置服务）+ `wows-ime.Pages`（Shell / HomePage / SettingsPage 页面与绑定模型）
- 游戏配置写入 `<游戏根目录>\bin\` 下所有数字版本目录的 `res_mods\ime_config.xml`；`8842736` 仅为版本示例，禁止硬编码为回退方案
- 简单状态存 `ApplicationData.Current.LocalSettings`（`Settings.SchemaVersion`、`Settings.Language`、`Game.SelectedPath`）；列表数据存 SQLite `settings.db`（`custom_game_paths`、`custom_input_methods`）；旧 `config.json` 启动时迁移并改名为 `.migrated`
- 发布启用 full 裁剪，`ILLink.Descriptors.xml` 保留 `WinRT.Runtime`，勿删
- 版本号在 `Package.appxmanifest`；发版流程：改版本号 → 更新 `CHANGELOG.md` → 运行 `tag.ps1` 打 tag 推送 → GitHub Action 自动创建 Release
- 本地测试按打包应用处理（MSIX），推荐使用 `wows-ime (Package)` 启动配置

## 常用命令

```powershell
# 构建（x64）
msbuild /t:build /restore /p:Configuration=Release /p:Platform=x64

# 发布（自包含 + full 裁剪）
dotnet publish -c Release -r win-x64
```
