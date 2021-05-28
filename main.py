from wechat_doc_spider import WechatSpider
from commons import create_table
#from app_operate import AppRobot
from save import totxt, one_toexcel


if __name__ == '__main__': 

    #control_type的类型：
    #wechat_history:查询指定公众号的所有历史文章
    #wechat_update: 更新指定公众号的最新文章（最多更新前50篇，也就是前十页的推送）
    #totxt: 把指定公众号的所有文章的标题和内容存入txt文档中，一篇文章一个文档
    #one_toexcel: 把指定公众号的所有文章的标题、时间、摘要、永久链接等存入一个excel文档中，一个公众号用一个worksheet,合成一个大文件
    control_type = 'wechat_update'

    #user_id是你的公众号的账号，注意是账号不是名字
    user_id = 'miawild46' #'miawild46' #'tjtmyt'
    #pwd是你的公众号的密码
    pwd = 'Myt1997224' #'Myt1997224' #'myt1997224'

    #biz列表里的是你要爬取的所有公众号中文名，不要改变顺序
    biz = ['中信证券研究','国泰君安证券研究','中泰证券研究','张忆东策略世界','股市荀策','分析师徐彪','招商银行研究','李迅雷金融与投资','XYSTRATEGY','华泰策略研究','伍戈经济笔记','姜超的投资视界','中金点睛','固收彬法','雪涛宏观笔记','戴康的策略世界','中金策略']
    #每一个公众号的信息都存入一个mysql表中，不要改变顺序
    table_name = ['zx','gt','zt','zyd','gsxc','xb','zs','lxl','xystrategy','ht','wg','jc','zjdj','gsbf','xt','dk','zjcl']
    #想要爬取的公众号序号： 以上面的顺序为编号(从0开始)，例：中信证券研究 target=0、国泰君安证券研究 target=1
    target = 16

    is_continue_wechat = True

    #以下是用来获得点赞、评论数的，需要更进一步的抓包工具，暂时用不上
    #chat_name = '文件传输助手'
    #is_continue_app = False
    #finish_null = True 

    start_date = "2021-05-01 00:00:00"
    end_date = "2021-05-24 00:00:00"

    #建立一个模板table——biz,每个公众号的表格都复制biz模板表格的结构
    create_table('biz')

    if control_type == 'wechat_history':
        i = biz[target]
        j = table_name[biz.index(i)]
        create_table(j)
        wechat = WechatSpider(user=user_id, pwd=pwd, biz=i, table=j, is_continue=is_continue_wechat)
        wechat.get_all_paper()
        #one_toexcel(biz=i,table=j) #自动把当前的公众号信息存入excel
    elif control_type == 'wechat_update':
        for i in biz[4:]:
            j = table_name[biz.index(i)]
            create_table(j)
            wechat1 = WechatSpider(user=user_id, pwd=pwd, biz=i, table=j, is_continue=False)
            wechat1.get_new_paper()
            one_toexcel(biz=i,table=j, start_date=start_date, end_date=end_date) #自动把当前的公众号信息存入excel
    elif control_type == 'one_toexcel': #手动把公众号信息存入excel
        i = biz[target]
        j = table_name[biz.index(i)]
        one_toexcel(biz=i,table=j, start_date=None, end_date=None)
    elif control_type == 'totxt':
        for i in biz:
            j = table_name[biz.index(i)]
            totxt(biz=i, table=j)
    #else: //这部分用来获取点赞和评论，暂时用不到这个功能
        #robot = AppRobot()
        #robot.run(biz_name=biz, chat_name=chat_name, is_continue=is_continue_app, finish_null=finish_null)
