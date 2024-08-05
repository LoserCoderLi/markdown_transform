# 注意：原开发环境是Linux
# 介绍
## 功能
- 上传内容：zip压缩包
- zip压缩包：markdown文件，存放照片的文件夹
- 格式（**文件名自取**）：
	- package.zip
		- test.md
		- imgs
			- image1
				- 1.png
				- ……
			- image2
				- 1.png
				- ……
			- ……

- 选择转化格式：HTML、PDF、word（**三选一**）
- 输出相应的文件

## 相关事项
1. 一个网页关闭后，你上传的文件也会跟着一起被删除
2. 每天的凌晨一点，会删除前一天的内容（考虑到和第一条有重叠的功能，可能会删除一个功能）
3. 转化为PDF或者word是，会要求输入格外内容。包括题目、作者、声明、左页眉、右页眉，对于PDF，还会要求输入封面的页脚
4. 转化为PDF或者word会还会生成目录（根据markdown文件的标题）
5. PDF的目录会直接显示出来，而对于word的目录，需要自己在打开docx文件的软件里面进行格外操作，下面我将进行举例说明

**软件：wps**
打开转化后的docx文件，向下翻，会发现”**Table of Contents**“字眼，点击后会发现有“**更新目录**”这个字眼，点击即可。

**点击**

![点击](/imgs/1Hc6Py8KRK9MFIvD.png)

**显示**

![输入图片说明](/imgs/G8fNnmMNI8BFkGOx.png)

<!--stackedit_data:
eyJoaXN0b3J5IjpbLTExNjI5NTkxNjZdfQ==
-->
