# VBA
Excel小工具

# <center>**Excel调用MD5宏使用文档**

### 一.说明

宏文件名称：md5.xla

MD5加密方式：32位加密，大写

调用加密函数：Md5_String_Calc()

### 二.加载步骤

1.打开Excel，点击左上角“文件”，然后选择“选项”
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\1.jpg)

2.点选“自定义功能区”
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\2.jpg)

3.勾选“开发者工具”
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\3.jpg)

4.选择功能栏中的“开发者工具”
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\4.jpg)

5.选择“Excel加载项”
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\5.jpg)

6.选择“浏览”
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\6.jpg)

7.找到加密文件“md5.xla”的位置，加载文件，确定
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\7.jpg)

### 三.使用功能

1.调用“Md5_String_Calc()”函数对需要加密的字符串加密
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\8.jpg)
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\9.jpg)
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\10.jpg)

2.小写加密，使用Excel自有的LOWER()函数,LOWER(Md5_String_Calc())
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\11.jpg)
![](C:\Users\Administrator\Desktop\SWG\MD5宏\ps\12.jpg)
