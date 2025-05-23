#!/bin/bash

name="tizk.py"

echo "TexcellentConverter ${name} running"

python src/${name}

if [ $? -eq 0 ]; then
    echo "正常に終了"
else
    echo "エラー．終了コード: $?"
fi 