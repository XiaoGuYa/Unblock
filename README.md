安装代理
以 管理员身份 打开 Powershell，Windows 10 快捷入口：Win + X - Windows Powershell(管理员)(A)，复制以下代码，右键粘贴到命令行回车，打开安装菜单。

Set-ExecutionPolicy -ExecutionPolicy Unrestricted -Force
Invoke-Expression -Command (Invoke-WebRequest -UseBasicParsing -Uri https://raw.githubusercontent.com/XiaoGuYa/Unblock/main/unm.ps1).Content


随后选择 1 即安装。安装完毕后选择 3 运行。如需添加开机自启，则执行 7。最后输入 0 退出。

----------------------------
     网易云解锁安装脚本
            V0.02 2019.09.12
                @AUTHOR LOGI
----------------------------
0. 退出
1. 安装
2. 卸载
3. 运行
4. 停止
5. 更新
6. 局域网共享
7. 添加开机自启
8. 取消开机自启
----------------------------
请选择:
