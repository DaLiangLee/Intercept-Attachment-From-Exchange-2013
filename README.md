# Intercept-Attachment-From-Exchange-2013
Intercept Attachment From Exchange 2013

安装

	1. 复制下面几个文件到Exchange server 上的文件夹C:\:
		1. InterceptAttachment.dll
		2. Install.ps1
	2. 用 Exchange Management Shell 运行 install.ps1脚本
	3. 关闭Exchange Management Shell
	4. 确保Microsoft Exchange Transport 服务已启动
	5. 截获的附件会放在C:\test 文件夹

卸载
	1. 用 Exchange Management Shell 运行 Uninstall.ps1
	2. 关闭Exchange Management Shell

