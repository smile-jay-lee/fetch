# 诊断脚本：HTTP 401 阻塞原因分析

## 问题背景

当前full run发现约65%的review URL返回HTTP 401，导致采标提取率为0。需要确定：
- **是否为IP封禁**：如果IP被目标服务器加入黑名单，单个请求也会被拒绝
- **是否为并发限制**：如果目标网站实施了速率限制，高并发会触发限制
- **是否为混合因素**：某些条件下限制更严格

## 诊断脚本原理

### 测试场景

脚本按照以下逐步升级的场景进行测试：

| 场景 | 工作线程 | 延迟 | 用途 |
|-----|--------|------|-----|
| 单线程无延迟 | 1 | 0ms | 基准测试，检测基本IP状态 |
| 单线程500ms延迟 | 1 | 500ms | 检测速率限制效果 |
| 单线程1s延迟 | 1 | 1000ms | 检测完全避免速率限制 |
| 多线程(3工作) | 3 | 0ms | 低并发测试 |
| 多线程(10工作) | 10 | 0ms | 当前默认配置 |
| 多线程(20工作) | 20 | 0ms | 高并发测试 |

### 诊断指标

对于每个场景，收集：
- HTTP状态码分布（200, 401, 其他）
- 401占比（%）
- 平均响应时间
- 总耗时

### 诊断判断逻辑

#### 判定为 IP 封禁
```
如果单线程无延迟的401率 > 50%
→ 即使单个请求也被拒绝
→ 问题在IP层面（服务器已识别该IP为恶意/机器人）
```

#### 判定为 并发限制
```
如果 多线程(20工作)401率 - 单线程无延迟401率 > 20%
→ 并发数越高，401率越高
→ 目标网站对并发有限制（防止DDoS）
```

#### 判定为 速率限制
```
如果 添加延迟后401率明显下降（>10%）
→ 请求速度太快被识别为异常
→ 添加延迟可以解决问题
```

## 使用方法

### 快速运行（30个采样URL）

```bash
cd d:\project\fetch\tests
python diagnose_blocking.py ..\部分.工作簿1.xlsx
```

### 完整运行（100个采样URL）

```bash
python diagnose_blocking.py ..\部分.工作簿1.xlsx 100
```

### 输出

脚本会输出：

1. **分析结果汇总表**
   ```
   测试场景                  配置                          200  401  401占比  平均响应  总耗时
   ───────────────────────────────────────────────────────────────────────────────────
   单线程无延迟            workers=1, delay=0ms            20   10   33.3%    0.45s    15.2s
   单线程500ms延迟         workers=1, delay=500ms          25   5    16.7%    0.52s    21.5s
   ...
   ```

2. **诊断结论**
   - 明确指出IP封禁/并发限制或混合原因
   - 给出对应的解决方案

3. **详细结果JSON**
   - 保存到 `diagnose_blocking_results.json`
   - 包含每个请求的状态码、响应时间等

## 预期输出示例

### 场景 A：IP 封禁
```
❌ 【结论】：极可能是 IP 封禁（或被加入黑名单）
   - 单线程无延迟的401率高达 65.0%
   - 即使单个请求也被拒绝，说明问题在服务器端已识别的IP

【对策】：
   1. 更换客户端IP（VPN/代理）
   2. 等待足够长时间（可能需要数小时到数天）
   3. 尝试使用不同的User-Agent和Headers
   4. 考虑使用住宅代理池或SOCKS代理
```

### 场景 B：并发限制
```
✓ 【结论】：主要是 并发控制触发（目标网站有速率限制）
   - 单线程401率: 20.0%
   - 20线程401率: 65.0%
   - 差异: 45.0%

【对策】：
   1. ⭐ 降低并发数（改为2-5个工作线程）
   2. 在请求间加延迟（500ms-1s）
   3. 实现指数退避重试机制
   4. 添加随机延迟避免规律性请求
```

### 场景 C：速率限制
```
✓ 【结论】：目标网站实施了 速率限制
   - 无延迟401率: 60.0%
   - 1s延迟401率: 15.0%
   - 延迟改善: 45.0%

【对策】：
   1. 在fetch.py中添加 --min-request-interval 参数（默认500ms）
   2. 实现自适应延迟（遇到401时自动增加延迟）
   3. 降低默认workers数（从10改为3-5）
```

## 对应的解决方案

### 如果诊断结果是 IP 封禁
1. **立即可行**：
   - 使用不同网络（4G卡、不同ISP的VPN）
   - 改变请求特征（改User-Agent、加随机Headers）

2. **短期（1-24小时）**：
   - 部署反向代理（nginx）用本地代理进行请求
   - 使用AWS/Azure等云服务IP进行请求

3. **长期**：
   - 部署代理池（免费：free-proxy-list，付费：Bright Data等）
   - 考虑直接联系网站管理员说明用途

### 如果诊断结果是 并发限制
1. **立即修改fetch.py**：
   ```python
   # 改变默认workers
   parser.add_argument('--workers', type=int, default=3)  # 从10改为3
   
   # 加入请求间延迟
   parser.add_argument('--min-interval', type=float, default=0.5)  # 500ms
   ```

2. **进行智能重试**：
   ```python
   # 对401进行重试，带指数退避
   def fetch_with_retry(url, max_retries=3):
       for attempt in range(max_retries):
           try:
               result = fetch(url)
               if result.status != 401:
                   return result
               if attempt < max_retries - 1:
                   time.sleep(2 ** attempt)  # 1s, 2s, 4s...
           except Exception as e:
               if attempt == max_retries - 1:
                   raise
   ```

3. **监控和调整**：
   - 记录每个场景的401率
   - 找到最优的workers数和延迟
   - 可能的最优值：workers=2-3, min_interval=1.0s

### 如果诊断结果是 速率限制
流程同上（并发限制），重点在于加延迟而不是降并发数。

## 脚本限制和注意事项

1. **采样量**：默认30个URL是为了快速诊断（约3-5分钟）
   - 对于更准确的结果，建议增加到100+

2. **网络波动**：单次诊断可能受网络质量影响
   - 建议在网络稳定时段运行
   - 可多次运行对比结果

3. **服务器状态**：目标网站的限制策略可能随时间变化
   - 诊断结果反映的是当前时刻的状态
   - 需要定期重新诊断

4. **URL特异性**：某些特定URL可能有特殊限制
   - 如果某个URL总是被拒，可能是业务逻辑问题
   - 详细结果JSON可查看每个URL的结果

## 后续步骤

1. **立即运行诊断**：
   ```bash
   python tests/diagnose_blocking.py 部分.工作簿1.xlsx 50
   ```

2. **根据诊断结果，修改fetch.py**（见下面的方案制定）

3. **验证修改效果**：
   - 运行小样本(100行)重新测试
   - 与诊断结果对比

## 相关修改方案

### 方案1：降低并发 + 加延迟
```python
# fetch.py 第675行
parser.add_argument('--workers', type=int, default=3)  # 降至3
parser.add_argument('--request-interval', type=float, default=0.5, 
                   help='请求间延迟(秒)')
```

### 方案2：智能重试机制
```python
# 在process_row_task()中添加retry逻辑
def fetch_with_retry(url, max_retries=3, timeout=10):
    for attempt in range(max_retries):
        try:
            ...
        except HTTPError as e:
            if e.code == 401 and attempt < max_retries - 1:
                wait_time = min(2 ** attempt, 10)  # 指数退避，最高10s
                logger.warning(f"收到401，{wait_time}s后重试")
                time.sleep(wait_time)
                continue
            raise
```

### 方案3：自适应并发
```python
# 根据401率动态调整workers
if http_401_rate > 50:
    workers = max(workers // 2, 1)  # 遇到高401率时自动降并发
```

---

下一步：请运行诊断脚本，我会根据结果制定具体的修改方案。
