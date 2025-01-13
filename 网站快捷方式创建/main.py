import os
import requests
from urllib.parse import urlparse
import winshell
from win32com.client import Dispatch
import re
from bs4 import BeautifulSoup
import time
from PIL import Image, ImageTk
import io
import tkinter as tk
from tkinter import messagebox

def get_site_title(url):
    """获取网站标题"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        response = requests.get(url, headers=headers, timeout=5)
        
        # 更好的编码处理
        content_type = response.headers.get('content-type', '').lower()
        if 'charset=' in content_type:
            encoding = content_type.split('charset=')[-1]
        else:
            encoding = response.apparent_encoding
        
        response.encoding = encoding
        
        soup = BeautifulSoup(response.text, 'html.parser')
        title = soup.title.string if soup.title else None
        
        if title:
            # 清理标题文本
            title = title.strip()
            # 移除常见的网站后缀
            title = re.sub(r'[-_](首页|官网|官方网站|网站|在线).*$', '', title)
            # 确保标题是有效的UTF-8字符串
            title = title.encode('utf-8', errors='ignore').decode('utf-8')
            # 移除文件名中的非法字符，保留中文
            title = re.sub(r'[<>:"/\\|?*\n\r\t]', '', title)
            # 如果标题太长，截取前30个字符（考虑中文字符）
            if len(title) > 30:
                title = title[:30].strip()
            # 确保文件名有效
            if not all(ord(c) < 128 or ('\u4e00' <= c <= '\u9fff') for c in title):
                title = ''.join(c for c in title if ord(c) < 128 or ('\u4e00' <= c <= '\u9fff'))
            return title if title.strip() else get_site_name(url)
        
        return get_site_name(url)
    except Exception as e:
        print(f"获取网站标题失败: {e}")
        return get_site_name(url)

def get_site_name(url):
    """获取网站名称（作为备用）"""
    try:
        parsed_url = urlparse(url)
        name = parsed_url.netloc.replace('www.', '')
        name = name.split('.')[0]
        name = ' '.join(word.capitalize() for word in name.split('-'))
        return name
    except:
        return "Website Shortcut"

def download_favicon(url):
    """下载网站的favicon图标并转换为ICO格式"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }
        
        parsed_url = urlparse(url)
        base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
        print(f"正在获取网站 {base_url} 的图标...")

        # 创建固定的图标保存目录
        icon_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'WebIcons')
        os.makedirs(icon_dir, exist_ok=True)

        # 为网站创建唯一的图标文件名
        site_hash = hash(base_url) & 0xffffffff
        png_icon_filename = f"icon_{site_hash}.png"
        ico_icon_filename = f"icon_{site_hash}.ico"
        png_icon_path = os.path.join(icon_dir, png_icon_filename)
        ico_icon_path = os.path.join(icon_dir, ico_icon_filename)

        # 如果PNG图标已存在且不为空，直接转换为ICO
        if os.path.exists(png_icon_path) and os.path.getsize(png_icon_path) > 0:
            print(f"找到缓存的PNG图标: {png_icon_path}")
            # 转换为ICO
            result = convert_png_to_ico(png_icon_path, ico_icon_path)
            if result:
                return result
            return download_default_icon()

        # 尝试直接获取favicon.ico
        try:
            favicon_url = f"{base_url}/favicon.ico"
            print(f"尝试从 {favicon_url} 下载图标...")
            response = requests.get(favicon_url, headers=headers, timeout=5)
            if response.status_code == 200 and len(response.content) > 0:
                with open(png_icon_path, 'wb') as f:
                    f.write(response.content)
                print(f"成功下载图标到: {png_icon_path}")
                # 转换为ICO
                result = convert_png_to_ico(png_icon_path, ico_icon_path)
                if result:
                    return result
                return download_default_icon()
            else:
                print("未找到直接的favicon.ico")
        except Exception as e:
            print(f"获取favicon.ico失败: {e}")

        # 尝试从HTML中查找图标
        try:
            print("尝试从网页HTML中查找图标链接...")
            response = requests.get(url, headers=headers, timeout=5)
            soup = BeautifulSoup(response.text, 'html.parser')
            
            # 按优先级排序的图标类型
            icon_rels = [
                'shortcut icon',
                'icon',
                'apple-touch-icon',
                'apple-touch-icon-precomposed'
            ]
            
            for icon_rel in icon_rels:
                for link in soup.find_all('link', rel=icon_rel):
                    icon_link = link.get('href')
                    if icon_link:
                        if icon_link.startswith('//'):
                            icon_link = parsed_url.scheme + ':' + icon_link
                        elif not icon_link.startswith(('http://', 'https://')):
                            if icon_link.startswith('/'):
                                icon_link = base_url + icon_link
                            else:
                                icon_link = base_url + '/' + icon_link
                        
                        print(f"尝试下载图标: {icon_link}")
                        try:
                            icon_response = requests.get(icon_link, headers=headers, timeout=5)
                            if icon_response.status_code == 200 and len(icon_response.content) > 0:
                                content_type = icon_response.headers.get('content-type', '').lower()
                                if any(t in content_type for t in ['image', 'icon', 'png', 'jpeg', 'gif']):
                                    with open(png_icon_path, 'wb') as f:
                                        f.write(icon_response.content)
                                    print(f"成功下载图标到: {png_icon_path}")
                                    # 转换为ICO
                                    result = convert_png_to_ico(png_icon_path, ico_icon_path)
                                    if result:
                                        return result
                                    return download_default_icon()
                                else:
                                    print(f"无效的图标类型: {content_type}")
                        except Exception as e:
                            print(f"下载图标失败: {e}")
                            continue

            print("在HTML中未找到有效的图标链接")
        except Exception as e:
            print(f"解析HTML失败: {e}")

        print("未能找到有效的网站图标，将使用默认图标")
        return download_default_icon()  # 使用默认图标

    except Exception as e:
        print(f"下载favicon过程中出错: {e}")
        return download_default_icon()  # 使用默认图标

def download_default_icon():
    """下载默认图标并返回其路径"""
    default_icon_url = "https://www.loliapi.com/acg/pp/"
    default_icon_path = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'WebIcons', 'defaultimg.png')
    
    try:
        response = requests.get(default_icon_url, timeout=5)
        if response.status_code == 200:
            with open(default_icon_path, 'wb') as f:
                f.write(response.content)
            print(f"成功下载默认图标到: {default_icon_path}")
            return default_icon_path
        else:
            print("下载默认图标失败，使用空图标")
            return None
    except Exception as e:
        print(f"下载默认图标时出错: {e}")
        return None

def convert_png_to_ico(png_path, ico_path):
    """将PNG图像转换为ICO格式"""
    try:
        with Image.open(png_path) as img:
            img.save(ico_path, format='ICO')
        print(f"成功将PNG图标转换为ICO: {ico_path}")
        return ico_path
    except Exception as e:
        print(f"转换PNG为ICO时出错: {e}")
        # 删除可能损坏的文件
        if os.path.exists(png_path):
            try:
                os.remove(png_path)
            except:
                pass
        if os.path.exists(ico_path):
            try:
                os.remove(ico_path)
            except:
                pass
        # 获取默认图标
        default_icon = download_default_icon()
        if default_icon:
            try:
                # 将默认图标转换为ICO
                with Image.open(default_icon) as img:
                    img.save(ico_path, format='ICO')
                return ico_path
            except:
                return default_icon
        return None

def create_shortcut(url, name=None):
    """创建Internet快捷方式并返回图标路径"""
    try:
        if not name:
            name = get_site_title(url)
            
        # 确保名称不为空且合法
        if not name or len(name.strip()) == 0:
            name = get_site_name(url)
            
        # 清理文件名
        name = name.strip()
        
        # 确保名称是有效的UTF-8字符串
        name = name.encode('utf-8', errors='ignore').decode('utf-8')
        
        # 只保留ASCII字符和中文字符
        name = ''.join(c for c in name if ord(c) < 128 or ('\u4e00' <= c <= '\u9fff'))
        
        # 移除文件名中的非法字符
        name = re.sub(r'[<>:"/\\|?*\n\r\t]', '', name)
        
        # 如果处理后的名称为空，使用默认名称
        if not name or len(name.strip()) == 0:
            name = "Website Shortcut"
            
        # 限制文件名长度（考虑中文字符）
        if len(name.encode('utf-8')) > 100:
            new_name = ''
            for c in name:
                if len((new_name + c).encode('utf-8')) > 100:
                    break
                new_name += c
            name = new_name.strip()
            
        # 确保文件名不以点或空格结尾
        name = name.rstrip('. ')

        # 获取桌面路径
        desktop = winshell.desktop()
        
        # 创建快捷方式路径，确保文件名唯一
        base_name = name
        counter = 1
        shortcut_path = os.path.join(desktop, f"{base_name}.url")
        while os.path.exists(shortcut_path):
            base_name = f"{name} ({counter})"
            shortcut_path = os.path.join(desktop, f"{base_name}.url")
            counter += 1
        
        print(f"准备创建快捷方式，文件名: {base_name}")
        
        # 下载favicon
        icon_path = download_favicon(url)
        
        # 创建.url文件
        with open(shortcut_path, 'w', encoding='utf-8') as f:
            f.write('[InternetShortcut]\n')
            f.write(f'URL={url}\n')
            
            if icon_path and os.path.exists(icon_path):
                # 使用特殊格式设置图标
                abs_icon_path = os.path.abspath(icon_path)
                f.write(f'IconIndex=0\n')
                f.write(f'IconFile={abs_icon_path}\n')
                # 添加HotKey
                f.write('HotKey=0\n')
                # 添加工作目录
                f.write(f'WorkingDirectory={os.path.dirname(abs_icon_path)}\n')
                # 添加ShowCommand
                f.write('ShowCommand=1\n')
                # 添加修改时间
                f.write(f'Modified={int(time.time())}\n')
            else:
                # 使用Chrome的默认图标
                chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                f.write(f'IconIndex=0\n')
                f.write(f'IconFile={chrome_path}\n')
            
            # 添加其他必要的属性
            f.write('[{000214A0-0000-0000-C000-000000000046}]\n')
            f.write('Prop3=19,11\n')
        
        print(f"成功创建快捷方式: {shortcut_path}")
        return shortcut_path, icon_path
        
    except Exception as e:
        print(f"创建快捷方式时出错: {e}")
        print(f"当前尝试创建的文件名: {name}")
        return None, None

def on_create_shortcut():
    """处理创建快捷方式的事件"""
    url = url_entry.get().strip()
    name = name_entry.get().strip() or None

    if not url.startswith(('http://', 'https://')):
        url = 'https://' + url

    shortcut_path, icon_path = create_shortcut(url, name)

    if shortcut_path:
        messagebox.showinfo("成功", f"成功创建快捷方式: {shortcut_path}")
        if icon_path:
            # 显示图标预览
            img = Image.open(icon_path)
            img.thumbnail((64, 64))  # 缩小图标
            icon_preview = ImageTk.PhotoImage(img)
            icon_label.config(image=icon_preview)
            icon_label.image = icon_preview  # 保持引用
    else:
        messagebox.showerror("错误", "创建快捷方式失败")

# 创建tkinter窗口
root = tk.Tk()
root.title("网站快捷方式创建器")

# 输入框和标签
tk.Label(root, text="请输入网站地址:").pack(pady=5)
url_entry = tk.Entry(root, width=50)
url_entry.pack(pady=5)

tk.Label(root, text="自定义快捷方式名称:").pack(pady=5)
name_entry = tk.Entry(root, width=50)
name_entry.pack(pady=5)

# 创建按钮
create_button = tk.Button(root, text="创建快捷方式", command=on_create_shortcut)
create_button.pack(pady=20)

# 图标预览标签
icon_label = tk.Label(root)
icon_label.pack(pady=10)

# 启动tkinter主循环
root.mainloop()
