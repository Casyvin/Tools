#!/bin/bash
# 清理并编译项目
echo "开始编译项目..."
mvn clean package

# 检查编译是否成功
if [ $? -ne 0 ]; then
    echo "Maven编译失败"
    exit 1
fi

# 创建输出目录
mkdir -p dist

# 获取实际的jar文件名
JAR_FILE=$(ls target/*.jar | grep -v sources | grep -v javadoc | head -n 1)
if [ -z "$JAR_FILE" ]; then
    echo "未找到可执行的jar文件"
    exit 1
fi

JAR_NAME=$(basename "$JAR_FILE")
echo "找到jar文件: $JAR_NAME"

# 在macOS上使用不同的方式获取JAVA_HOME
if [[ "$OSTYPE" == "darwin"* ]]; then
    JAVA_HOME_PATH=$(/usr/libexec/java_home)
else
    JAVA_HOME_PATH=$(dirname $(dirname $(readlink -f $(which java))))
fi

echo "Java路径: $JAVA_HOME_PATH"

# macOS打包（包含Java运行时）
echo "开始打包..."
jpackage --input target/ \
  --name PLTools \
  --app-version 1.0 \
  --main-class com.sws4cloud.pltools.PLToolsApplication \
  --main-jar "$JAR_NAME" \
  --type dmg \
  --runtime-image "$JAVA_HOME_PATH" \
  --dest dist \
  --mac-package-signing-prefix com.sws4cloud.pltools

# 检查打包结果
if [ $? -eq 0 ]; then
    echo "打包完成，检查dist目录:"
    ls -la dist/
else
    echo "打包失败，请检查错误信息"
fi
