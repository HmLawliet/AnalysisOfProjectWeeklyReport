# 项目简介
该项目是从阿里邮箱爬取项目文件并解析生成png图片，再通过钉钉发送给指定人的程序

钉钉的设置,请联系钉钉客服，
钉钉后台创建一个app,获取accessKey
通过钉钉开发的sdk说明 
通过创建的这个app来实现工作通知的

# 下载pyecharts
pip install pyecharts

# 下载 snapshot_phantomjs
# 首先要安装phantomjs
npm install phantomjs
pip install snapshot_phantomjs

# 安装PIL 系统centos7
# 安装epel源
yum -y install epel-release
# 对pip进行升级
pip install --upgrade pip
# 安装缺少类库及PIL
yum -y install python-devel
yum install libjpeg libjpeg-devel zlib zlib-devel freetype freetype-devel lcms lcms-devel
yum -y install python-imaging

pip install pillow



# 下载html转化image的软件
# 从http://wkhtmltopdf.org 下载wkhtmltoimage
# weekly.py 模块的 drawTable()
# config = imgkit.config(wkhtmltoimage='wkhtmltoimage的安装目录')


# 找到macro文件 一般在  Lib/site-packages/pyecharts/render/templates
# 找到 {% if chart._component_type == "table" %}
# 修改 
.fl-table {
                margin: 20px;    -----> 修改成 margin: auto; 表格居中显示 
                border-radius: 5px;
                font-size: 12px;
                border: none;
                border-collapse: collapse;
                max-width: 100%;
                white-space: nowrap;
                word-break: keep-all;
            }

