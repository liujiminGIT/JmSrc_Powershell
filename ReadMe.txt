■Github DesktopのProxy追加方法

 C:\Users\JimmyLiu\.gitconfig に以下のエントリが追加される。

 [http]
	proxy = proxy.foo.com:80
 [https]
	proxy = proxy.foo.com:80

proxy=[username:password]@xxx.xxx.xxx.xxx:8080


■Github DesktopのコマンドラインGit

C:\Users\JimmyLiu\AppData\Local\GitHubDesktop\app-1.6.5\resources\app\git\cmd


■Pushエラー

fatal: unable to access　　　schannel: next InitializeSecurityContext failed: >Unknown error (0x80092012) 
対処方法

　git config --global http.sslVerify false