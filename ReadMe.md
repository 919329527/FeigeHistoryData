#  说明：

照例转了个`.exe`文件出来：  https://fuwenyue.lanzoup.com/ijx8G00ca3ab

#### 一、用途

[飞鸽](https://fxg.jinritemai.com/ffa/mshop/homepage/index) 是抖音店铺的后台管理系统，脚本用于抓取(店铺&客服)：

- 昨日数据
- 当月数据（1号至昨天）
- 周数据（昨天-6至昨天）

- 数据合并

![](https://gitee.com/fuwenyue/tuchuang/raw/master/utools/16452752394191645275239350.png)

#### 二 、用法

登陆账号进入店铺，获取Cookie后粘贴至`cookie.txt`，每行一个。同一账号下的不同店铺用的是不同Cookie。以下两种格式均可，自动忽略字符数小于1000的行。

```markdown
passport_csrf_token_**********
passport_csrf_token_**********
```

```mark
xxx手机旗舰店
passport_csrf_token_**********
xxx官方旗舰店
passport_csrf_token_**********
```

#### 三、Pyinstaller

`Pyinstaller -F -i favicon.ico feige.py`