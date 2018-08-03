# QQ-Crawler
a crawler to get qq friends' emtions

1.所有爬取的信息全部放在一个excel文件中

> 这个文件是提前建立好的 并且需要手动添加两张表单以及两张表单的表头
>
> 分别命名为friend_info 和 Black 
>
> 两张表的表头A1,B1位置为 名字 和 号码

2.先运行getQQfriends.py 获取所有好友的名字和qq号码 并且存放于friend_info中

3.再运行getFriendsEmotion.py 爬取说说

