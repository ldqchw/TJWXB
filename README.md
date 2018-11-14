# TJWXB
<h1>1、Web_Check</h1>
<ul>
<li>1.1、使用selenium进行截图
<li>1.2、使用tldextract进行格式化提取域名
<li>1.3、使用pyttsx当完成程序后语音提示
<li>1.4、使用logging进行日志记录
</ul>
<h2>使用方法</h2>
<p>python Dwszb_Check.py</p>
<p>安装所需要的库PS:Windows运行pyttsx需要安装pywin32</p>
<h1>2、Template</h1>
<ul>
<li>2.1、使用figlet生成banner
<li>2.2、使用getopt命令行输入
<li>2.3、使用xlwings读写Excel
<li>2.4、使用win32com进行Word批量替换
</ul>
<h2>使用方法</h2>
<p>python XQZG-Template.py -h </p>
<p>安装所需要的库</p>
<h1>3、SSH</h1>
<ul>
<li>3.1、paramiko 实现SSH登录
<li>3.2、xlrd、xlutils读写excel
<li>3.3、Mem_Free 可用内存/G
<li>3.4、Disk_Waring 磁盘是否超过80%
<li>3.5、CPU_Waring CPU系统是否超过80%
</ul>
<h2>使用方法</h2>
<p>python ssh.py </p>
<p>安装所需要的库</p>
<p>远程登录到服务器然后执行命令最后将结果存储到excel</p>
<h1>4、URL-Verification</h1>
<ul>
<li>1、读取execl第3个sheets页（H2:H1359）
<li>2、判断域名是否带www，如果带www就去掉，不带保留subdomain
<li>3、将格式化的数据保存在I列中，另存为：d:\URL.xlsx
</ul>
<h2>使用方法</h2>
<p>python URL-Verification.py </p>
<p>安装所需要的库</p>
<p>可以直接修改读取保存列表H2:H1359和'I'+ str(i)</p>

**PS:   
统一安装所需依赖库  
cd TJWXB  
pip  install -r requirements.txt**

notepad：  
自动导出所需依赖库  
pip freeze > requirements.txt


 
