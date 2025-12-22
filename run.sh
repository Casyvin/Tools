#!/bin/bash
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
java -cp "$DIR/*:$DIR/libs/*" com.sws4cloud.pltools.PLToolsApplication
