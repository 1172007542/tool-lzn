#!/bin/bash

# 设置内存使用率阈值
MEMORY_THRESHOLD=70

# 获取内存使用率
MEMORY_USAGE=$(free | grep Mem | awk '{printf "%.2f", $3/$2 * 100}')

# 检查内存使用率是否超过阈值
if (( $(echo "$MEMORY_USAGE > $MEMORY_THRESHOLD" | bc -l) )); then
    # 清理缓存
    sync
    echo 3 > /proc/sys/vm/drop_caches
    
    # 释放交换空间
    swapoff -a
    swapon -a
    
    # 记录日志
    echo "$(date): 内存使用率 $MEMORY_USAGE%, 已清理" >> /var/log/memory_monitor.log
fi
