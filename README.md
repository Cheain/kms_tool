##KMS激活工具

**本地激活服务器使用SystemRage的开源项目[py-kms](https://github.com/SystemRage/py-kms)，本项目的主要工作为增加易用性，实现图形化界面**


改动源项目中timezones.py中_detect_timezone_php()函数
```Python
if tomatch == (tz._tzname, -tz._utcoffset.seconds, indst):
```
修改为
```Python
if (tomatch[1], tomatch[2]) == (-tz._utcoffset.seconds, indst):
```
原因是在中文的Windows系统中`time.tzname[0]`返回`中国标准时区`而不是英文系统中的`CST`无法做比较，得不出任何时区结果，该函数在Linux中可以正常工作。



运行`main.py`功能：
- 开启本机KMS服务器：在本机搭建KMS服务器，监听0.0.0.0:1688端口，为本机或局域网内计算机提供KMS服务器的功能
- 可手动输入KMS激活服务器地址，默认本机地址127.0.0.1，也可输入局域网内（运行本脚本KMS服务器的电脑IP）或者互联网上可用的KMS服务器地址
- 安装GVLK密钥：当本机Windows未安装GVLK密钥时不能通过KMS的方式激活，需要先安装密钥；软件内置Windows Vista 到Windos 10 、Windows Server 2008 到Windows Server 2016 的各版本GVLK密钥，从微软MSDN资料库取得
- 激活Windows：自动从输入的KMS服务器处激活Windows，请确保KMS服务器可用
- 搜索Office位置：从注册表读取Windows安装位置，如果有安装多个版本的Office再次点击可切换下一个版本，如果搜索到的安装位置有误可以手动输入位置
- 激活Office：自动从输入的KMS服务器处、搜索或输入的Office安装处激活Office；Office 2013 及以后的版本激活不支持本机建立的KMS服务器，只能通过局域网内或互联网上等非本机的KMS服务器来激活Office

未完成功能：
- pyinstaller打包后按钮命令不执行，直接运行脚本可以执行
- 搜索功能在Windows 10 中无效， Windows 7 有效

打包命令`pyinstaller -F -w -i kms_icon.ico main.py`