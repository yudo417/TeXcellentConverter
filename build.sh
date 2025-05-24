#!/bin/bash

echo "TexcellentConverter dev mode running"

python src/app_new.py

if [ $? -eq 0 ]; then
    echo "正常に終了"
else
    echo "エラー．終了コード: $?"
fi 