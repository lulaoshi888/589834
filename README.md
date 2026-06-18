# 奇偶分析工具 - APK 构建指南

通过 GitHub Actions 在云端自动构建 APK，无需本地安装任何工具。

## 构建方法

### 方法一：使用 GitHub Actions（推荐，无需本地 Linux）

1. **注册 GitHub 账号**（如果没有）：https://github.com/signup

2. **创建新仓库**：
   - 点击右上角 `+` → `New repository`
   - 仓库名随意，例如 `parity-apk`
   - 设为 **Private**（私有仓库）
   - 点击 `Create repository`

3. **上传文件**：
   ```
   上传以下 3 个文件到仓库根目录：
   ├── main.py              # 主程序
   ├── buildozer.spec       # 构建配置
   └── .github/
       └── workflows/
           └── build-apk.yml  # 自动构建工作流
   ```

4. **自动构建**：
   - 上传完成后，进入仓库的 `Actions` 标签页
   - 会看到一个名为 `Build APK` 的工作流正在运行
   - 构建约 **30-40 分钟**（首次需要下载 Android SDK）

5. **下载 APK**：
   - 构建完成后，点击工作流进入详情
   - 在 `Artifacts` 部分下载 `parity-analysis-apk.zip`
   - 解压得到 `judgeparity-1.0.0-arm64-v8a-debug.apk`

### 方法二：手动构建（需要 Linux 环境）

```bash
# 在 Ubuntu/Debian 上执行
sudo apt update
sudo apt install -y git zip unzip openjdk-17-jdk python3-pip autoconf \
    libtool pkg-config zlib1g-dev libncurses5-dev libncursesw5-dev \
    libtinfo5 cmake libffi-dev libssl-dev libltdl-dev

pip3 install --user buildozer
buildozer android debug
# APK 在 bin/ 目录下
```

## 使用说明

1. 在手机上安装 APK
2. 将数据文件（.txt 或 .xlsx）放入手机存储
3. 打开 App，点击 `📁 选择数据文件` 选择文件
4. 或点击 `✏️ 手动输入` 追加最新数据
5. App 自动分析最后 8 行数据并显示预测结果
