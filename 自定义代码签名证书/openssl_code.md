前期准备工作：安装openssl



Microsoft Windows [版本 10.0.19043.2364]
(c) Microsoft Corporation。保留所有权利。

**openssl genrsa -aes128 -out aa.key 2048**

Enter PEM pass phrase:

Verifying - Enter PEM pass phrase:

aa可换成其他字符，加密方式可以是-des3，没有硬性要求

<u>Enter PEM pass phrase</u> 代表私钥加密密码，需要输入两次

------

**openssl req -new -x509 -key aa.key -days 3650 -sha1 -out ab.cer**

Enter pass phrase for aa.key:
You are about to be asked to enter information that will be incorporated into your certificate request.
What you are about to enter is what is called a Distinguished Name or a DN.
There are quite a few fields but you can leave some blank.
For some fields there will be a default value,If you enter '.', the field will be left blank.

Country Name (2 letter code) [AU]:

State or Province Name (full name) [Some-State]:

Locality Name (eg, city) []:

Organization Name (eg, company) [Internet Widgits Pty Ltd]:

Organizational Unit Name (eg, section) []:

Common Name (e.g. server FQDN or YOUR name) []:

Email Address []:

-key后需指定key文件，-sha1参数为必须，因为Excel不支持sha1以外的签名算法，ab可换成其他字符串，但在之后的步骤中需要拼写正确

<u>Enter pass phrase for aa.key</u> 代表输入引入私钥key文件的密码

------

**openssl pkcs12 -export -name XX -in ab.cer -inkey aa.key -out ae.pfx**

Enter pass phrase for aa.key:

Enter Export Password:

Verifying - Enter Export Password:

此步骤为合并公钥cer文件和私钥key文件为pfx证书，XX、ae可更换为其他字符串，-name可不指定

<u>Enter pass phrase for aa.key:</u> 需输入私钥密码

<u>Enter Export Password:</u> 需设置pfx证书密码并二次确认 (证书导入时需要此密码进行验证)，此密码不可与私钥密码相同



后续操作：

cmd中输入certmgr后回车启动证书管理器

选择存储路径->个人->证书，导入pfx证书，输入pfx证书密码，导入选项仅选择"包括所有扩展属性"，证书存储至个人，导入成功后即刻生效

带宏的Excel文档中，选择开发工具->Visual Basic，菜单栏中选择工具->数字签名，选择需要签署的证书，确定后保存文件，再次打开文件启用宏时将额外提示签名信息