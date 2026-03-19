"""
测试具体URL：查看为什么这样的review URL无法获取数据
"""

import sys
from urllib.request import urlopen, Request
from urllib.error import HTTPError

# 测试用户提供的URL
test_url = "https://std.samr.gov.cn/gb/search/gbDetailed?id=234D7936AB54E194E06397BE0A0AA0A9&review=true"

print("=" * 70)
print("测试具体URL")
print("=" * 70)
print(f"\nURL: {test_url}\n")

# 尝试1：最简单的请求（无headers）
print("【尝试1】最简单的请求（无headers）")
try:
    resp = urlopen(test_url, timeout=5)
    body = resp.read()
    print(f"  ✓ HTTP {resp.status}")
    print(f"  ✓ Content-Length: {len(body)}")
    if len(body) > 0:
        print(f"  ✓ 成功获取内容")
except HTTPError as e:
    print(f"  ✗ HTTP {e.code}")
except Exception as e:
    print(f"  ✗ {type(e).__name__}: {e}")

# 尝试2：带基础User-Agent
print("\n【尝试2】带基础User-Agent")
headers = {
    'User-Agent': 'Mozilla/5.0'
}
try:
    req = Request(test_url, headers=headers)
    resp = urlopen(req, timeout=5)
    body = resp.read()
    print(f"  ✓ HTTP {resp.status}")
    print(f"  ✓ Content-Length: {len(body)}")
except HTTPError as e:
    print(f"  ✗ HTTP {e.code}")
except Exception as e:
    print(f"  ✗ {type(e).__name__}: {e}")

# 尝试3：带完整User-Agent（Chrome）
print("\n【尝试3】带完整User-Agent（Chrome）")
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
}
try:
    req = Request(test_url, headers=headers)
    resp = urlopen(req, timeout=5)
    body = resp.read()
    print(f"  ✓ HTTP {resp.status}")
    print(f"  ✓ Content-Length: {len(body)}")
    if len(body) > 0 and len(body) < 100:
        print(f"  内容: {body.decode('utf-8', errors='replace')}")
except HTTPError as e:
    print(f"  ✗ HTTP {e.code}")
except Exception as e:
    print(f"  ✗ {type(e).__name__}: {e}")

# 尝试4：带Referer (从std.samr.gov.cn来)
print("\n【尝试4】带Referer和完整headers")
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Referer': 'https://std.samr.gov.cn/',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.9',
    'Accept-Encoding': 'gzip, deflate',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
}
try:
    req = Request(test_url, headers=headers)
    resp = urlopen(req, timeout=5)
    body = resp.read()
    print(f"  ✓ HTTP {resp.status}")
    print(f"  ✓ Content-Length: {len(body)}")
    if len(body) > 0:
        print(f"  ✓ 成功获取内容")
        # 检查内容
        content = body.decode('utf-8', errors='replace')
        if '采标' in content:
            print(f"  ✓ 页面包含'采标'")
        if '标准' in content:
            print(f"  ✓ 页面包含'标准'")
except HTTPError as e:
    print(f"  ✗ HTTP {e.code}")
    # 检查是否是401
    if e.code == 401:
        print(f"  说明：被认为是未授权的请求")
except Exception as e:
    print(f"  ✗ {type(e).__name__}: {e}")

# 尝试5：不带&review=true参数
print("\n【尝试5】不带&review=true（去掉参数）")
url_without_review = "https://std.samr.gov.cn/gb/search/gbDetailed?id=234D7936AB54E194E06397BE0A0AA0A9"
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
    'Referer': 'https://std.samr.gov.cn/',
}
try:
    req = Request(url_without_review, headers=headers)
    resp = urlopen(req, timeout=5)
    body = resp.read()
    print(f"  ✓ HTTP {resp.status}")
    print(f"  ✓ Content-Length: {len(body)}")
except HTTPError as e:
    print(f"  ✗ HTTP {e.code}")
except Exception as e:
    print(f"  ✗ {type(e).__name__}: {e}")

print("\n" + "=" * 70)
print("分析")
print("=" * 70)
print("""
std.samr.gov.cn/gb/search/gbDetailed 是查看标准详情的页面

&review=true 参数可能表示：
  • review=true: 显示暗稿/审查版本（可能需要权限）
  • 无参数或review=false: 显示公开版本

问题可能是：
  1. &review=true 需要登录身份
  2. 该端点有特殊的访问限制
  3. 只有特定的User-Agent才能访问
""")
