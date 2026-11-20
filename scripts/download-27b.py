# -*- coding: utf-8 -*-
"""下载 Qwen3.8-27B 高端模型（unsloth 动态量化 Q3_K_XL，13.2GB）到 client-base models/"""
from modelscope import snapshot_download

model_dir = snapshot_download(
    'unsloth/Qwen3.8-27B-GGUF',
    allow_file_pattern=['Qwen3.8-27B-UD-Q3_K_XL.gguf'],
    local_dir='models',
)
print('DONE ->', model_dir)
