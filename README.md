# resume-export
从Word或者PDF简历中导出电话和邮箱的Python脚本

## 使用

```shell

cd resume-export
python -m venv venv
.\venv\Scripts\Activate.ps1
pip install -r requirements.txt
python start.py

```
## 注意

- **首次运行根据提示将需要导出的简历文件放入脚本生成的origin目录下，以后每次运行替换origin中的文件即可。**

- **错误文件保存在error目录下**

- **除了origin目录外其余文件每次运行脚本都会清除，注意保存**
