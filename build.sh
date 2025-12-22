
#!/bin/bash

echo "开始打包PL-TOOLS项目..."

# 1. 编译项目
echo "编译项目..."
mvn clean package

if [ $? -ne 0 ]; then
    echo "编译失败，退出打包"
    exit 1
fi

# 2. 创建输出目录
mkdir -p dist/windows dist/macos

# 3. 获取主JAR文件
JAR_FILE=$(ls target/*.jar | grep -v sources | grep -v javadoc | head -n 1)
if [ -z "$JAR_FILE" ]; then
    echo "未找到主JAR文件"
    exit 1
fi

JAR_NAME=$(basename "$JAR_FILE")
echo "找到主JAR文件: $JAR_NAME"

# 4. Windows打包
echo "开始Windows打包..."
jpackage --input target/ \
  --name PLTools \
  --app-version 1.0 \
  --main-class com.sws4cloud.pltools.PLToolsApplication \
  --main-jar "$JAR_NAME" \
  --type exe \
  --win-shortcut \
  --win-menu \
  --dest dist/windows

# 5. macOS打包
echo "开始macOS打包..."
jpackage --input target/ \
  --name PLTools \
  --app-version 1.0 \
  --main-class com.sws4cloud.pltools.PLToolsApplication \
  --main-jar "$JAR_NAME" \
  --type dmg \
  --mac-package-name "PLTools" \
  --dest dist/macos

echo "打包完成！"
echo "Windows版本位于: dist/windows/"
echo "macOS版本位于: dist/macos/"