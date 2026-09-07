# -*- coding: utf-8 -*-
"""
合并 seen_announcements.json 专用脚本，供 workflow 调用。

把"远端最新版本"(origin/main 分支上的内容) 和 "本地这一轮新增的记录"
做并集，写到 /tmp/merged_seen.json。

两组"已见公告"取并集永远是安全操作，不会因为两次运行前后脚碰上
而互相覆盖、丢失任何一边新增的记录——这样就不用依赖git的文本行级
合并（那种方式一旦两边都改了同一处内容就容易冲突报错）。
"""
import json
import subprocess

LOCAL_PATH = "seen_announcements.json"
OUTPUT_PATH = "/tmp/merged_seen.json"

try:
    with open(LOCAL_PATH, "r", encoding="utf-8") as f:
        local_data = set(json.load(f))
except Exception:
    local_data = set()

try:
    remote_raw = subprocess.check_output(
        ["git", "show", "origin/main:seen_announcements.json"]
    )
    remote_data = set(json.loads(remote_raw))
except Exception:
    remote_data = set()

merged = sorted(local_data | remote_data)

with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
    json.dump(merged, f, ensure_ascii=False, indent=2)

print(f"合并完成：本地{len(local_data)}条 + 远端{len(remote_data)}条 -> 去重后{len(merged)}条")
