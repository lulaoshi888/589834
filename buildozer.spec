[app]

title = 奇偶分析工具
package.name = judgeparity
package.domain = org.parity
source.dir = .
source.include_exts = py,png,jpg,kv,atlas
version = 1.0.0

android.permissions = INTERNET, READ_EXTERNAL_STORAGE, WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 21
android.use_androidx = True
android.accept_sdk_license = True
android.log_level = 2

requirements = python3,kivy==2.2.1,openpyxl,xlrd,plyer

orientation = portrait
fullscreen = 1
android.screen_sizes = small, normal, large, xlarge
android.archs = arm64-v8a

[buildozer]
log_level = 2
