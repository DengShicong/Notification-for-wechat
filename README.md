# Notification-for-wechatNotification-for-wechat

![Static Badge](https://img.shields.io/badge/netmiko-4.2.0-blue%20) ![Static Badge](https://img.shields.io/badge/openpyxl-3.1.2-green%20) ![Static Badge](https://img.shields.io/badge/requests-2.31.0-red) ![Static Badge](https://img.shields.io/badge/openpyxl-3.1.2-yellow) ![Static Badge](https://img.shields.io/badge/XlsxWriter-3.1.9-oringo) ![Static Badge](https://img.shields.io/badge/paramiko-3.3.1-pink) ![Static Badge](https://img.shields.io/badge/python-3.10.6-9cf)



## 简介

该脚本是集成版本具有以下功能：

1. 自动化巡检通过CLI交互获取回显信息。
2. 将回显信息通过WxPusher发送到公众号，点击查看详细内容。
3. 将回显信息制作简易巡检报告，设备health信息与温度信息以高亮颜色设置（红/绿），表示处于正常与异常
4. sftp下载交换机配置文件vcboot.cfg/boot.cfg，对应AOS 6.07.0/8.0。
5. 将配置文件与巡检报告打包发送邮件到运维团队，支持多人接收。



## 效果

 简易巡检报告：![image](https://github.com/DengShicong/Notification-for-wechat/blob/main/images/5bdb0705f552d83b7d1b3afad1f1425.png)


微信告警通知：

![image](https://github.com/DengShicong/Notification-for-wechat/blob/main/images/2b43df42e83059ffc972b05ee7ac356.png)

 告警推送内详细内容：

![image](https://github.com/DengShicong/Notification-for-wechat/blob/main/images/2bcfc2c0316ba59d28e30dcd7cd6fd0.png)

邮件通知：
![image](https://github.com/DengShicong/Notification-for-wechat/blob/main/images/b622ba18ada68775ccb205a4872a3e0.png)



## Usage

脚本导入pycharm可直接使用，将自定义参数如Token，uid，邮箱之类替换即可。
与template.xlsx模板文件同一目录下，在模板内设置多个巡检设备。

Token获取与公众号相关信息可查看WxPusher文档：https://wxpusher.zjiecode.com/docs/#/
