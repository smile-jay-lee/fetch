#!/usr/bin/env python3
"""
验证改进的headers是否能够解决采标提取问题
"""

import sys
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.error import HTTPError
from urllib.parse import urlparse, urlencode, parse_qsl, urlunparse

# 导入改进的build_request_headers函数
sys.path.insert(0, str(Path(__file__).parent.parent))
from fetch import build_request_headers

def test_newgbinfo_endpoint():
    """测试newGbInfo端点"""
    print("=" * 60)
    print("测试1: 新的GbInfo端点")
    print("=" * 60)
    
    url = "https://openstd.samr.gov.cn/bzgk/gb/newGbInfo"
    params = {
        "limit": 10,
        "offset": 0,
        "filters": '[{"name":"ics","value":""},{"name":"classify","value":""},{"name":"keyword","value":""}]',
    }
    full_url = f"{url}?{urlencode(params)}"
    
    try:
        req = Request(full_url, headers=build_request_headers(full_url))
        print(f"\n📝 请求URL: {full_url[:80]}...")
        print(f"\n📋 请求Headers:")
        for k, v in build_request_headers(full_url).items():
            if k == "User-Agent":
                print(f"   {k}: {v[:50]}...")
            elif len(str(v)) > 50:
                print(f"   {k}: {str(v)[:50]}...")
            else:
                print(f"   {k}: {v}")
        
        with urlopen(req, timeout=10) as resp:
            data = resp.read()
            status = resp.status
            content_length = len(data)
            
            print(f"\n✓ 响应状态: HTTP {status}")
            print(f"✓ 响应大小: {content_length} 字节")
            
            # 检查是否包含标准数据
            if content_length > 1000:
                print(f"✓ 数据充足 (> 1KB)")
                # 尝试解析JSON看是否包含数据
                if b'"data"' in data:
                    print(f"✓ 包含data字段")
                data_text = data.decode("utf-8", errors="ignore")
                if '"name"' in data_text or "标准" in data_text:
                    print(f"✓ 包含标准名称")
                return True
            else:
                print(f"⚠ 数据不足")
                return False
                
    except HTTPError as e:
        print(f"✗ HTTP错误 {e.code}: {e.reason}")
        return False
    except Exception as e:
        print(f"✗ 错误: {type(e).__name__}: {e}")
        return False

def test_review_endpoint():
    """测试review端点"""
    print("\n" + "=" * 60)
    print("测试2: Review端点（用户的具体URL）")
    print("=" * 60)
    
    # 使用用户提供的具体URL
    url = "https://std.samr.gov.cn/gb/search/gbDetailed"
    params = {
        "id": "234D7936AB54E194E06397BE0A0AA0A9",
        "review": "true"
    }
    full_url = f"{url}?{urlencode(params)}"
    
    try:
        req = Request(full_url, headers=build_request_headers(full_url))
        print(f"\n📝 请求URL: {full_url[:80]}...")
        print(f"\n📋 请求Headers (关键的):")
        headers = build_request_headers(full_url)
        print(f"   User-Agent: {headers['User-Agent'][:50]}...")
        print(f"   Referer: {headers.get('Referer', 'N/A')}")
        print(f"   Accept: {headers.get('Accept', 'N/A')[:50]}...")
        
        with urlopen(req, timeout=10) as resp:
            data = resp.read()
            status = resp.status
            content_length = len(data)
            
            print(f"\n✓ 响应状态: HTTP {status}")
            print(f"✓ 响应大小: {content_length} 字节")
            
            # 检查是否包含标准数据
            data_text = data.decode("utf-8", errors="ignore")
            keywords = ["采标", "标准", "编号", "发布"]
            found_keywords = [kw for kw in keywords if kw in data_text]

            if found_keywords:
                print(f"✓ 包含中文内容: {found_keywords}")
            
            if content_length > 1000:
                print(f"✓ 数据充足 (> 1KB)")
                return True
            else:
                print(f"⚠ 数据不足")
                return False
                
    except HTTPError as e:
        print(f"✗ HTTP错误 {e.code}: {e.reason}")
        return False
    except Exception as e:
        print(f"✗ 错误: {type(e).__name__}: {e}")
        return False

def main():
    print("\n🔧 测试改进的headers对采标提取的影响\n")
    
    test1_result = test_newgbinfo_endpoint()
    test2_result = test_review_endpoint()
    
    print("\n" + "=" * 60)
    print("📊 测试结果总结")
    print("=" * 60)
    print(f"newGbInfo端点: {'✓ 通过' if test1_result else '✗ 失败'}")
    print(f"Review端点: {'✓ 通过' if test2_result else '✗ 失败'}")
    
    if test1_result and test2_result:
        print("\n✅ 所有测试通过！headers修复有效")
        print("📈 下一步: 运行完整数据集测试采标提取率\n")
        return 0
    else:
        print("\n⚠️ 部分测试失败，可能需要进一步调整\n")
        return 1

if __name__ == "__main__":
    sys.exit(main())
