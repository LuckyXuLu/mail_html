# coding:utf-8   #强制使用utf-8编码格式

# 用Python 在 AWS 调用邮件接口：
import smtplib
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from multiprocessing import Process, Queue
import random
import xlrd
import csv
import re
import time


class SendEmail:
    def __init__(self, dwell_time, int_initnum, Abnormal_through, csv_name_list, meil_list):
        self.csv_name_list = csv_name_list
        self.dwell_time = dwell_time
        self.int_initnum = int_initnum
        self.Abnormal_through = Abnormal_through
        self.cov_str = None
        self.proxies = []
        self.read_name_mail()

        print(self.cov_str)
        # 土木邮件标题与发件人
        if "土木" in self.cov_str:
            self.theme = ["EI_Call for Papers_ICCE2018", "ICCE2018 Call for Papers---EI Indexed",
                          "Call for Papers--The 5th International Conference on Civil Engineering (EI Indexed)"]

            self.meil_list = meil_list
        # 地质邮件标题与发件人
        elif "地质" in self.cov_str:
            self.theme = ["EI_Call for Papers_ICGRMSD2018", "ICGRMSD2018 Call for Papers---EI Indexed",
                          "Call for Papers--6th International Conference on Geology Resources Management and Sustainable Development (EI Indexed)"]

            self.meil_list = meil_list
        elif "计算机" in self.cov_str:
            self.theme = ["Invitation Letter from AICS2019"]

            self.meil_list = meil_list
        elif "材料" in self.cov_str:
            self.theme = ["Invitation Letter from MEMS2019"]

            self.meil_list = meil_list

        self.f = 1
        print("总邮箱: {}个".format(len(self.proxies)))
        print("等待10秒....")
        time.sleep(10)

    def ctrl_function(self, cov_str, row_list):
        with open("{}.txt".format(cov_str), "a") as a_3:
            a_3.write("{}*{}\n".format(row_list[0], row_list[1]))

    def read_name_mail(self):
        for csv_name in self.csv_name_list:
            try:
                workbook = xlrd.open_workbook("{}.csv".format(csv_name))
            except:
                pass
            else:
                self.cov_str = csv_name
                break

        if not self.cov_str:
            print("请确认文件名")
            return

        table = workbook.sheets()[0]
        for row in range(table.nrows):
            row_list = table.row_values(row)
            if len(row_list) == 2:
                if row_list[1][-2:].isalpha() and re.match(
                        '^[A-Za-z0-9](([A-Za-z0-9\-\_]{1,}(\.[A-Za-z0-9][A-Za-z0-9\-\_]{0,}){0,})|([A-Za-z0-9\-\_]{0,}(\.[A-Za-z0-9][A-Za-z0-9\-\_]{0,}){1,}))@(((([A-Za-z][A-Za-z\-\_]{1,}(\.[A-Za-z][A-Za-z\-\_]{1,}){0,})|([A-Za-z][A-Za-z\-\_]{0,}(\.[A-Za-z][A-Za-z\-\_]{1,}){1,})|(163)|(uc3m)|(189)|(126)|(263)|(univ-lyon1)|(139)|(2k)|(csp2)|(uniroma1)|(uniroma3)|(mining3)|(mk3)|(unina2)|(jasatirta1))\.(([A-z]{2,3})|(info)|(coop))$)|(vip\.163\.com$))',
                        row_list[1]):  # 判断字符串是否为字母
                    if row_list[1][-9:] == "gmail.com":
                        self.proxies.append((row_list[0], row_list[1]))
                    else:
                        a_list = ["outlook.com", "firstname", "yahoo", "mail.ru", "hotmail", "gmail.co",
                                  "griffith.edu.au"]
                        if [i for i in a_list if i in row_list[1]]:
                            self.ctrl_function(self.cov_str + "拒绝域名", row_list)
                            continue

                        self.proxies.append((row_list[0], row_list[1]))
                else:
                    po = [".jp"]
                    if [i for i in po if i == row_list[1][-len(i):]]:
                        self.proxies.append((row_list[0], row_list[1]))
                        continue

                    if self.Abnormal_through:
                        self.proxies.append((row_list[0], row_list[1]))
                    else:
                        self.ctrl_function(self.cov_str + "奇怪字符邮件", row_list)
                        continue
            else:
                self.ctrl_function(self.cov_str + "错误", row_list)

    # 土木英文
    def civil_english_email(self, name):
        BODY_HTML = """<!DOCTYPE html>
                        <html lang="en">
                        <head>
                            <meta charset="UTF-8">
                            <title>EI_Call for Papers_ICCE2018</title>
                        </head>
                        <body>
                        <div class="div_01" style="width: 600px; margin: 10px auto; font-family: Times New Roman; background-color: #cccccc;">
                            <p class="p_02" style="margin: 0px auto;"><img src="cid:image1" style="width: 600px; height:322px"></p>
                            <h2 class="h2_04" style="text-align: center; margin: 0px auto;">5th International Conference on Civil Engineering (ICCE2018)</h2>
                            <h2 class="h2_05" style="text-align: center; margin: 0px auto;">Call for Papers---EI Indexed</h2>
                            <div class="div_03" style="margin: 0px 5px;">
                                <h4 style="margin: 0px; font-weight: lighter;">Dear %s,</h4>
                                <!--<h4 style="margin: 8px 0px 8px; font-weight: lighter;">-->
                                    <!--It is our great pleasure to invite you to join in and submit your lastest research results to 5th-->
                                    <!--International Conference on Civil Engineering (ICCE2018) on Dec. 20-21, 2018, which will be held in Nanchang-->
                                    <!--, one of the most famous national historic and cultural cities in China.-->
                                <!--</h4>-->

                                <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                    &#8195;&#8195;It is our great pleasure to invite you to submit your newest research to 5th International Conference on Civil Engineering
                                    (ICCE2018) on Dec. 20-21, 2018, which will be held in Nanchang, one of the most famous national historic and cultural
                                    cities in China.
                                </h4>

                                <h4 style="margin: 0px 0px 8px; font-weight: lighter;">
                                    &#8195;&#8195;ICCE2018 is organized by Nanchang Institute of Technology, Civil Engineering Academy of Jiangxi Province and
                                    Hubei Zhongke Institute of Geology and Environment Technology, co-organized by journal of Rock and Soil
                                    Mechanics.
                                </h4>

                                <h4 style="margin: 0px; font-weight: lighter;">
                                    &#8195;&#8195;ICCE2018 offers a platform for experts, scholars and engineering staff in the field of civil engineering to
                                    make exchanges and seeks high-quality, original papers that address the theory, design, development and
                                    evaluation of ideas, tools, techniques and methodologies in (but not limited to):
                                </h4>
                                <h4 style="margin: 0px; font-weight: lighter;">1. Structural Engineering;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">2. Construction Engineering;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">3. Geotechnical Engineering;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">4. Architecture and Building Materials;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">5. Construction Materials: Traditional and Advanced;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">6. Roads and Bridges;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">7. Tunnel Subway and Underground Engineering;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">8. Water Resources Engineering Architecture and Urban Planning;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">9. Other Related Topics.</h4>

                                <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                    &#8195;&#8195;All contributions to the meeting will be peer-reviewed by Academic Board, and the review period will be
                                    about one month. Part of the excellent contributions to the ICCE will be published by EI journals, such as
                                    Journal of Rock Mechanics and Geotechnical Engineering, Rock and Soil Mechanics, Magazine of Civil
                                    Engineering, Electronic Journal of Structural Engineering.
                                </h4>

                                <!--<h4 style="margin: 0px 0px 8px; font-weight: lighter;">-->
                                    <!--ICCE2018 also want to invite some people to be our guest editors and academic board members. If you are-->
                                    <!--interested in it, you can send me your curriculum vitae as well.-->
                                <!--</h4>-->

                                <h4 style="margin: 0px; font-weight: lighter;">Important dates:</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">Paper Submission Deadline: Dec. 5, 2018</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">Notification of Acceptance/Rejection: Dec. 10, 2018</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">Conference Date: Dec. 20-21, 2018</h4>

                                <h4 style="margin: 8px 0px 8px; font-weight: lighter;">Please visit our website
                                    <a href="http://icce.easyace.cn/">http://icce.easyace.cn/</a>
                                    for more details.
                                </h4>

                                <h4 style="margin: 0px 0px 8px; font-weight: lighter;">
                                    Should you have any questions or concerns feel free to contact us. (E-mail:
                                    <a href="mailto: iccecn@163.com?subject=EI_Call for Papers_ICCE2018">iccecn@163.com</a>)
                                </h4>

                                <h4 style="margin: 0px; font-weight: lighter;"><a href="https://www.linkedin.com/in/changbo-cheng-1666ab16a/">Changbo CHENG in Linkedin</a></h4>
                                <h4 style="margin: 0px; font-weight: lighter;">Organizing Committee</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">ICCE2018</h4>

                                <h4 style="margin: 8px 0px 0px; font-weight: lighter;">If you do not want to receive this email again, please click
                                    <a href="mailto:icce@easyaca.cn?subject=This is an unsubscription request email!&body=This is an unsubscription request email! Please send it to me to unsubscribe!">here</a>
                                    for unsubscription!
                                </h4>
                            </div>
                        </div>
                        </body>
                """ % name

        return BODY_HTML

    # 土木中文
    def civil_chinese_email(self, name):
        BODY_HTML = """<!DOCTYPE html>
                                <html lang="en">
                                <head>
                                    <meta charset="UTF-8">
                                    <title>EI_Call for Papers_ICCE2018</title>
                                </head>
                                <body>
                                <div class="div_01" style="width: 600px; margin: 10px auto; font-family: Times New Roman; background-color: #cccccc;">
                                    <p class="p_02" style="margin: 0px auto;"><img src="cid:image1" style="width: 600px; height:322px"></p>
                                    <h2 class="h2_04" style="text-align: center; margin: 0px auto;">第五届土木工程国际学术会议</h2>
                                    <h2 class="h2_05" style="text-align: center; margin: 0px auto;">征文通知---EI全文检索</h2>
                                    <div class="div_03" style="margin: 0px 5px;">
                                        <h4 style="margin: 0px; font-weight: lighter;">尊敬的%s老师：</h4>
                                        <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                            <p style="text-indent:28px">您好！由南昌工程学院、江西省土木建筑学会、湖北省众科地质与环境技术服务中心主办，《岩土力学》期刊部、江西省水利土木工程基础设施安全重点实验室、江西省水利土木特种加固与安全监控工程研究中心协办的“2018·第五届土木工程国际学术会议”（ICCE 2018）将于2018年12月20¬-21日在中国南昌召开。大会诚邀国际和国内相关学科科技工作者与工程技术人员撰稿，积极参会，以期提高土木工程领域科研理论能力和工程实践水平，促进学科的发展和重大工程技术难题的解决。所有会议录用的论文，将作为会议文献公开出版并被Ei Compendex检索；优秀论文将推荐到相关EI期刊并被EI Compendex 和Scopus数据库全文检索。
                                        </h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">一、主办单位</h4>
                                        <h4 style="margin: 0px 0px 0px; font-weight: lighter;">南昌工程学院</h4>
                                        <h4 style="margin: 0px 0px 0px; font-weight: lighter;">江西省土木建筑学会</h4>
                                        <h4 style="margin: 0px 0px 0px; font-weight: lighter;">湖北省众科地质与环境技术服务中心</h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">二、协办单位</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">《岩土力学》期刊部</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">江西省水利土木工程基础设施安全重点实验室</h4>
                                        <h4 style="margin: 0px ; font-weight: lighter;">江西省水利土木特种加固与安全监控工程研究中心</h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">三、组织机构</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">
                                            &#8195;&#8195;大会设立大会学术委员会、组委会、大会秘书处等机构具体负责大会相关组织工作，将邀请相关领域的知名专家在会议上作主题报告。
                                        </h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">四、大会主题</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">创新驱动土木工程</h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">五、相关议题</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">01 建筑工程 02 岩土工程 03 结构工程 04 材料工程 05 工程勘察</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">06 道路和桥梁工程 07 水资源工程 08 地质工程 09 测量工程</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">10 防灾减灾工程 11 施工技术和装配式建筑施工(BIM)</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">12 隧道、地铁及地下工程 13 绿色建筑与环境保护</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">14 土木工程材料与新工艺、新方法、新技术</h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">六、会议出版物</h4>
                                        <h4 style="margin: 0px 0px 0px 0px; font-weight: lighter;">
                                            &emsp;&emsp;所有会议录用的论文，将作为会议文献公开出版并被Ei compendex检索；优秀论文将推荐到《Geotechnical Engineering》、《岩石力学与岩土工程学报（英文版）》、《岩土力学》等期刊并被EI Compendex 和Scopus数据库全文检索。同时，为了提高论文的被引用度，专刊将被发行到世界150个高校与研究机构。
                                        </h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">七、投稿方式</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">
                                            全文作者可通过官方邮箱：<a href="mailto: iccecn@163.com?subject=征文通知--第五届土木工程国际学术会议-- EI全文检索">iccecn@163.com</a>投稿，也可通过网站在线投稿。
                                        </h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">八、格式要求</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">
                                            请严格按格式样张排版：格式样张（<a href="http://icce.easyace.cn/uploadfile/ueditor/file/20180210/1518239645115170.doc">可点击下载</a>）
                                        </h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">九、重要期限</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">2018年12月5日截稿         2018年12月10日发完录用通知</h4>
                                        <h4 style="margin: 0px; font-weight: lighter;">2018年12月15日截止会议注册  2018年12月18日截止收费</h4>

                                        <h4 style="margin: 8px 0px 0px; font-weight: bold;">十、大会秘书处</h4>
                                        <h4 style="margin: 0px 0px 0px; font-weight: lighter;">湖北省众科地质与环境技术服务中心</h4>
                                        <h4 style="margin: 0px 0px 0px; font-weight: lighter;">联系人：李老师 张老师</h4>
                                        <h4 style="margin: 0px 0px 0px; font-weight: lighter;">电 话：027-87332517</h4>
                                        <h4 style="margin: 0px 0px 8px; font-weight: lighter;">地 址：湖北省武汉市珞狮路147号未来城A座2703室</h4>

                                        <h4 style="margin: 0px 0px 0px; font-weight: lighter;">
                                            附件1：征文通知（
                                            <a href="http://icce.easyace.cn/uploadfile/ueditor/file/20180820/1534752983931853.pdf">点击下载</a>）
                                        </h4>
                                        <h4 style="margin: 0px 0px 8px; font-weight: lighter;">
                                            附件2：格式样张（
                                            <a href="http://icce.easyace.cn/uploadfile/ueditor/file/20180210/1518239645115170.doc">点击下载</a>
                                            ）
                                        </h4>

                                        <h4 style="margin: 0px 0px 8px; font-weight: lighter;">
                                            如果您不想再接收此类邮件，请
                                            <a href="mailto:iccecn@163.com?subject=征文通知--第五届土木工程国际学术会议-- EI全文检索, 退订!&body=不想再接收此类邮件，退订!">点击这里</a>
                                            退订！
                                        </h4>

                                    </div>
                                </div>
                                </body>
                        """ % name

        return BODY_HTML

    # 地质英文
    def geology_english_email(self, name):
        BODY_HTML = """<!DOCTYPE html>
                        <html lang="en">
                        <head>
                            <meta charset="UTF-8">
                            <title>EI_Call for Papers_ICGRMSD2018</title>
                        </head>
                        <body>
                        <div class="div_01" style="width: 600px; margin: 10px auto; font-family: Times New Roman; background-color: #cccccc;">
                            <p class="p_02" style="margin: 0px auto;"><img src="cid:image1" style="width: 600px; height:322px"></p>
                            <h2 class="h2_04" style="text-align: center; margin: 0px auto;">6th International Conference on Geology Resources Management and Sustainable Development</h2>
                            <h2 class="h2_05" style="text-align: center; margin: 0px auto;">Call for Papers---EI Indexed</h2>
                            <div class="div_03" style="margin: 0px 5px; text-align: justify;">
                                <h4 style="margin: 0px; font-weight: lighter;">Dear %s,</h4>
                                <!--<h4 style="margin: 8px 0px 8px; font-weight: lighter;">-->
                                    <!--It is our great pleasure to invite you to join in and submit your lastest research results to 5th-->
                                    <!--International Conference on Civil Engineering (ICCE2018) on Dec. 20-21, 2018, which will be held in Nanchang-->
                                    <!--, one of the most famous national historic and cultural cities in China.-->
                                <!--</h4>-->

                                <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                    &#8195;&#8195;It is our great pleasure to invite you to submit your newest research
                                    to 6th International Conference on Geology Resources Management and Sustainable
                                    Development (ICGRMSD2018) on Dec. 27-28, 2018, which will be held in Beijing, China.
                                </h4>

                                <h4 style="margin: 0px 0px 8px; font-weight: lighter;">
                                    &#8195;&#8195;Organized by Hubei Zhongke Institute of Geology and Environment Technology, co-organized by School of Energy Resources, China university of Geosciences (Beijing), ICGRMSD2018 is to implement the harmonious development idea of China’s "One Belt and One Road" and ecological environment, give full play to the important role of geological science and technology in safeguarding economic and social development and promoting ecological progress in the new era, and finally promote the theory and technology innovation of engineering geology.
                                </h4>

                                <h4 style="margin: 0px; font-weight: lighter;">
                                    &#8195;&#8195;ICGRMSD2018 offers a platform for experts, scholars and engineering staff in the related field to make exchanges and seeks high-quality, original papers that address the theory, design, development and evaluation of ideas, tools, techniques and methodologies in (but not limited to):
                                </h4>
                                <h4 style="margin: 0px; font-weight: lighter;">1. Management of Geological and Mineral Resources and Sustainable Development;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">2. Safety Engineering and Evaluation of Geological Resources;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">3. Geological Resource Management and Environmental Economic Evaluation;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">4. Land Resource Management and Sustainable Development;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">5. Development, Planning and Management of Tourism Resources;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">6. The Value Evaluation of Geological Relics, Mining Parks and Geoparks;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">7. Mine Geological Environment Protection, Restoration and Control and Mine Geological Hazard Prevention;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">8. Investment and Operation Management of Mining Enterprises;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">9. The Development and Management of the Resource Industry;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">10. Investment and Operation Management of Geological Prospecting Units and Mining Enterprises;</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">11. Other Topics Related.</h4>

                                <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                    &#8195;&#8195;All contributions to ICGRMSD2018 will be published as conference proceedings by Aussino Academic Publishing House and submitted for EI Compendex (all the precious five ICGRMSD proceedings have been indexed in EI Compendex). Part of the excellent contributions to ICGRMSD2018 will be recommended to be published in Rock Mechanics and Rock Engineering, World of Mining - Surface and Underground, Journal of Disaster Research, etc., which will be indexed in SCI, EI and Scopus for full-text retrieval.
                                </h4>

                                <!--<h4 style="margin: 0px 0px 8px; font-weight: lighter;">-->
                                    <!--ICGRMSD2018 also want to invite scholars and experts to be our reviewers or editorial board members. If you are interested in it, you can send me your curriculum vitae as well.-->
                                <!--</h4>-->

                                <h4 style="margin: 0px; font-weight: lighter;">Important dates:</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">Paper Submission Deadline: Dec. 15, 2018</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">Notification of Acceptance/Rejection: Dec. 20, 2018</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">Conference Date: Dec. 27-28, 2018</h4>

                                <h4 style="margin: 8px 0px 8px; font-weight: lighter;">Please visit our website
                                    <a href="http://icgrmsd.easyace.cn/">http://icgrmsd.easyace.cn/</a>
                                    for more details.
                                </h4>

                                <h4 style="margin: 0px 0px 8px; font-weight: lighter;">
                                    Should you have any questions or concerns feel free to contact us.  (E-mail:
                                    <a href="mailto: icgrmsd2015@163.com?subject=EI_Call for Papers_ICGRMSD2018">icgrmsd2015@163.com</a>)
                                </h4>

                                <!--<h4 style="margin: 0px; font-weight: lighter;"><a href="https://www.linkedin.com/in/changbo-cheng-1666ab16a/">Changbo CHENG in Linkedin</a></h4>-->
                                <h4 style="margin: 0px; font-weight: lighter;">Organizing Committee</h4>
                                <h4 style="margin: 0px; font-weight: lighter;">ICGRMSD2018</h4>

                                <h4 style="margin: 8px 0px 0px; font-weight: lighter;">If you do not want to receive this email again, please click
                                    <a href="mailto:icgrmsd2018@easyaca.cn?subject=This is an unsubscription request email!&body=This is an unsubscription request email! Please send it to me to unsubscribe!">here</a>
                                    for unsubscription!
                                </h4>
                            </div>
                        </div>
                        </body>
                """ % name

        return BODY_HTML

    # 地质中文
    def geology_chinese_email(self, name):
        BODY_HTML = """<!DOCTYPE html>
                            <html lang="en">
                            <head>
                                <meta charset="UTF-8">
                                <title>EI_Call for Papers_ICGRMSD2018</title>
                            </head>
                            <body>
                            <div class="div_01" style="width: 600px; margin: 10px auto; font-family: Times New Roman; background-color: #cccccc;">
                                <p class="p_02" style="margin: 0px auto;"><img src="cid:image1" style="width: 600px; height:322px"></p>

                                <div class="div_03" style="margin: 0px 5px; text-align: justify;">
                                    <h4 style="margin: 0px; font-weight: lighter;">尊敬的 %s:</h4>
                                    <!--<h4 style="margin: 8px 0px 8px; font-weight: lighter;">-->
                                        <!--It is our great pleasure to invite you to join in and submit your lastest research results to 5th-->
                                        <!--International Conference on Civil Engineering (ICCE2018) on Dec. 20-21, 2018, which will be held in Nanchang-->
                                        <!--, one of the most famous national historic and cultural cities in China.-->
                                    <!--</h4>-->


                                    <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                        &#8195;&#8195;您好！由湖北省众科地质与环境技术服务中心主办，中国地质大学（北京）能源学院协办的“2018·第六届地质资源管理与可持续发展国际学术会议”（ICGRMSD2018）将于2018年12月27-28日在中国北京召开。大会诚邀国际和国内相关学科科技工作者与工程技术人员撰稿，积极参会，以期提高地质资源领域科研理论能力和工程实践水平，促进学科的发展和重大工程技术难题的解决。所有会议录用的论文，将作为会议文献公开出版并提交Ei compendex检索（前五届论文集均已被EI收录）；优秀论文将推荐到相关EI、SCI期刊。
                                    </h4>

                                    <h4 style="margin: 0px 0px 0px; font-weight: lighter;">
                                        附件：征文通知（
                                        <a href="http://icgrmsd.easyace.cn/uploadfile/ueditor/file/20180910/1536569542122313.pdf">点击下载</a>）
                                    </h4>
                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">一、主办单位</h4>
                                    <h4 style="margin: 0px 0px 0px; font-weight: lighter;">湖北省众科地质与环境技术服务中心</h4>

                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">二、协办单位</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">中国地质大学（北京）能源学院</h4>

                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">三、组织机构</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">
                                        &#8195;&#8195;大会设立大会学术委员会、组委会、大会秘书处等机构具体负责大会相关组织工作，将邀请相关领域的知名专家在会议上作主题报告。
                                    </h4>

                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">四、大会主题</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">地质资源管理与可持续发展</h4>

                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">五、相关议题</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">（1）采矿科学与技术 （2）地质矿产资源管理与可持续发展</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">（3）地质资源安全工程与评价  （4）地质资源管理与环境经济评价</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">（5）土地资源管理与可持续发展  （6）旅游资源开发、规划与管理</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">（7）地质遗迹、矿山公园和地质公园的价值评价</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">（8）矿山地质环境保护、恢复治理与矿山地质灾害防治</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">（9）矿山企业投资、营运管理  （10）资源产业发展与管理</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">（11）地勘单位、矿山企业投资与营运管理  （12）其它相关主题</h4>

                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">六、会议出版物</h4>
                                    <h4 style="margin: 0px 0px 0px 0px; font-weight: lighter;">
                                        &emsp;&emsp;所有会议录用的论文，将作为会议文献公开出版并被Ei compendex检索（前五届论文集均已被EI收录）；优秀论文将推荐到《Rock Mechanics and Rock Engineering（岩石力学和岩石工程）》、《World of Mining - Surface and Underground （采矿世界-地面和地下）》、《Journal of Disaster Research（灾害研究杂志）》等期刊以正刊形式出版，并提交SCI、EI Compendex、Scopus 数据库全文检索。同时，为了提高论文的被引用度，专刊将被发行到世界150个高校与研究机构。
                                    </h4>

                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">七、投稿方式</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">
                                        全文作者可通过官方邮箱：<a href="mailto:icgrmsd@easyaca.cn?subject=征文通知--第六届地质资源管理与可持续发展国际学术会议">icgrmsd@easyaca.cn</a>投稿，也可通过网站在线投稿。
                                    </h4>

                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">八、格式要求</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">
                                        请严格按格式样张排版：格式样张（<a href="http://icce.easyace.cn/uploadfile/ueditor/file/20180210/1518239645115170.doc">可点击下载</a>）
                                    </h4>

                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">九、重要期限</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">2018年12月15日截稿         2018年12月20日发完录用通知</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">2018年12月22日截止会议注册  2018年12月24日截止收费</h4>

                                    <h4 style="margin: 8px 0px 0px; font-weight: bold;">十、大会秘书处</h4>
                                    <h4 style="margin: 0px 0px 0px; font-weight: lighter;">湖北省众科地质与环境技术服务中心</h4>
                                    <h4 style="margin: 0px 0px 0px; font-weight: lighter;">联系人：李老师 张老师</h4>
                                    <h4 style="margin: 0px 0px 0px; font-weight: lighter;">电 话：027-87332517</h4>
                                    <h4 style="margin: 0px 0px 8px; font-weight: lighter;">地 址：湖北省武汉市珞狮路147号未来城A座2703室</h4>

                                    <h4 style="margin: 0px 0px 8px; font-weight: lighter;">
                                        如果您不想再接收此类邮件，请
                                        <a href="mailto:icgrmsd@easyaca.cn?subject=征文通知--第六届地质资源管理与可持续发展国际学术会议, 退订!&body=不想再接收此类邮件，退订!">点击这里</a>
                                        退订！
                                    </h4>
                                </div>
                            </div>
                            </body>
                    """ % name

        return BODY_HTML

    # 计算机英文评审
    def Computer_english_evaluation(self, name):
        BODY_HTML = """<!DOCTYPE html>
                            <html lang="en">
                            <head>
                                <meta charset="UTF-8">
                                <title></title>
                            </head>
                            <body>
                            <div class="div_01" style="width: 600px; margin: 10px auto; font-family: Times New Roman; background-color: #cccccc;">
                                <!--<p class="p_02" style="margin: 0px auto;"><img src="cid:image1" style="width: 600px; height:322px"></p>-->
                                <h2 class="h2_04" style="text-align: center; margin: 0px auto;">2019 International Conference on Artificial Intelligence and Computer Science (AICS2019)</h2>
                                <h1 class="h2_05" style="text-align: center; margin: 0px auto;">Invitation Letter</h1>
                                <h2 class="h2_05" style="text-align: center; margin: 0px auto;">for Scientific Committee Members and Guest Editors</h2>
                                <div class="div_03" style="margin: 0px 5px; text-align: justify;">
                                    <h4 style="margin: 0px; font-weight: lighter;">Dear %s,</h4>

                                    <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                        &#8195;&#8195;It is my honor to contact you. This is Dr. Cheng from Organizing Committee of 2019 International Conference on Artificial Intelligence and Computer Science (AICS2019)!
                                    </h4>

                                    <h4 style="margin: 0px 0px 8px; font-weight: lighter;">
                                        &#8195;&#8195;In view of your achievements and contribution in academic researches, the influence on related research areas, AICS2019 intends to invite you to join our Scientific Committee Members and Guest Editors Team, and act as our conferences academic guider promoting the development of the scientific fields jointly.
                                    </h4>

                                    <h4 style="margin: 0px; font-weight: lighter;">
                                        &#8195;&#8195;We are here sincerely inviting you to Scientific Committee Members and Guest Editors Team because of your great achievements and prestige in your areas.
                                    </h4>

                                    <h4 style="margin: 0px; font-weight: lighter;">
                                        &#8195;&#8195;If you feel interested in joining us, please fill in the application form below and send it to our mailbox <a href="mailto:easyace_cn@163.com">(easyace_cn@163.com)</a>. Your application will be handled in about 5 working days.
                                    </h4>

                                    <div style="width: 580px; height:680px; margin: 0px auto; border:#000000 solid 2px;">
                                        <div style="width: 578px; height:38px; margin: 0px auto; text-align:center; font-weight:bold; line-height:38px; border-bottom:1px solid;">Application Form</div>

                                        <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; font-weight:bold; line-height:30px; border-right:1px solid;">Name</div>
                                        <div style="width: 300px; height:31px; float:left; margin: 0px;"></div>
                                        <div style="width: 148px; height:30px; float:right; margin: 0px; border-left:1px solid;"></div>
                                        <div style="clear: both"></div>

                                        <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; font-weight:bold; line-height:30px; border-top:1px solid; border-right:1px solid; border-bottom:1px solid;">Email</div>
                                        <div style="width: 300px; height:31px; float:left; margin: 0px; border-top:1px solid; border-bottom:1px solid;"></div>
                                        <div style="width: 148px; height:30px; float:right; margin: 0px; border-left:1px solid;"></div>
                                        <div style="clear: both"></div>

                                        <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; line-height:30px; font-weight:bold; border-right:1px solid;">Title</div>
                                        <div style="width: 300px; height:31px; float:left; margin: 0px;"></div>
                                        <div style="width: 148px; height:30px; float:right; margin: 0px; border-left:1px solid;"></div>
                                        <div style="clear: both"></div>

                                        <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; line-height:30px; font-weight:bold; border-right:1px solid; border-top:1px solid; border-bottom:1px solid;">Institution</div>
                                        <div style="width: 300px; height:31px; float:left; margin: 0px; border-top:1px solid; border-bottom:1px solid;"></div>
                                        <div style="width: 148px; height:30px; float:right; margin: 0px; border-left:1px solid; border-bottom:1px solid;"></div>
                                        <div style="clear: both"></div>

                                        <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; font-weight:bold; line-height:30px; border-right:1px solid;">Research Fields</div>
                                        <div style="width: 449px; height:30px; float:left; margin: 0px;"></div>
                                        <div style="clear: both"></div>

                                        <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; line-height:30px; font-weight:bold; border-top:1px solid; border-bottom:1px solid; border-right:1px solid;">Address</div>
                                        <div style="width: 449px; height:30px; float:left; margin: 0px; border-top:1px solid; border-bottom:1px solid;"></div>
                                        <div style="clear: both"></div>

                                        <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; line-height:30px; font-weight:bold; border-right:1px solid;">Choose to be</div>
                                        <div style="width: 449px; height:30px; float:left; margin: 0px; text-align:center; line-height:30px; font-weight:bold; ">Scientific Committee Member □&#8195;&#8195;Guest Editor □&#8195;&#8195;Either □</div>
                                        <div style="clear: both"></div>

                                        <div style="width: 578px; height:38px; margin: 0px auto; text-align:center; line-height:38px; border-top:1px solid; font-weight:bold; border-bottom:1px solid;">Publication List</div>
                                        <div style="width: 578px; height:100px; margin: 0px auto; text-align:center;"></div>
                                        <div style="width: 578px; height:38px; margin: 0px auto; text-align:center; line-height:38px; border-top:1px solid; font-weight:bold; border-bottom:1px solid;">Achievements in Conferences and Journals </div>
                                        <div style="width: 578px; height:100px; margin: 0px auto; text-align:center;"></div>
                                        <div style="width: 578px; height:38px; margin: 0px auto; text-align:center; line-height:38px; border-top:1px solid; font-weight:bold; border-bottom:1px solid;">Other</div>
                                        <div style="width: 578px; height:100px; margin: 0px auto; text-align:center;"></div>
                                    </div>

                                    <!--<table style="width: 580px; height:600px; margin: 0px auto;">-->
                                        <!--<tr>-->
                                            <!--<td style="width: 578px; height:38px; margin: 0px auto; text-align:center; line-height:38px; border:#000000 solid 1px;">Application form</td>-->
                                        <!--</tr>-->
                                        <!--<tr>-->
                                            <!--<td style="width: 128px; height:30px; margin: 0px; padding: 0px; text-align:center; line-height:30px; border-left:1px solid; border-right:1px solid;">Name</td>-->
                                            <!--<td style="width: 300px; height:30px; margin: 0px; padding: 0px;"></td>-->
                                            <!--<td style="width: 148px; height:30px; margin: 0px; padding: 0px;border-left:1px solid; border-right:1px solid;"></td>-->
                                        <!--</tr>-->
                                        <!--<tr>-->
                                            <!--<td style="width: 128px; height:30px; margin: 0px; padding: 0px; text-align:center; line-height:30px; border:#000000 solid 1px;">Email</td>-->
                                            <!--<td style="width: 300px; height:30px; margin: 0px; padding: 0px; border-top:1px solid; border-bottom:1px solid;"></td>-->
                                            <!--<td style="width: 148px; height:30px; margin: 0px; padding: 0px; border-left:1px solid; border-right:1px solid;"></td>-->
                                        <!--</tr>-->
                                        <!--<tr>-->
                                            <!--<td>Title</td>-->
                                            <!--<td></td>-->
                                            <!--<td></td>-->
                                        <!--</tr>-->
                                        <!--<tr>-->
                                            <!--<td>Institution</td>-->
                                            <!--<td></td>-->
                                            <!--<td></td>-->
                                        <!--</tr>-->
                                        <!--<tr>-->
                                            <!--<td>Research Fields</td>-->
                                            <!--<td></td>-->
                                            <!--<td></td>-->
                                        <!--</tr>-->
                                        <!--<tr>-->
                                            <!--<td>Address</td>-->
                                            <!--<td></td>-->
                                            <!--<td></td>-->
                                        <!--</tr>-->
                                        <!--<tr>-->
                                            <!--<td>Choose to be</td>-->
                                            <!--<td></td>-->
                                            <!--<td></td>-->
                                        <!--</tr>-->
                                        <!--<tr>-->
                                            <!--<td>Institution</td>-->
                                            <!--<td>Scientific Committee Member □    Guest Editor □   Either □</td>-->
                                        <!--</tr>-->

                                    <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                        &#8195;&#8195;And we are looking for international cooperations with institutes, colleges, agencies in related areas and holding the conference together; if you feel it is the chance for us to work closely in academic communication, please do recommend AICS2019 to your organization.
                                    </h4>

                                    <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                        &#8195;&#8195;If you have any problems or suggestions on the cooperation, please feel free to contact me!
                                    </h4>

                                    <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                        &#8195;&#8195;AICS2019 welcomes the cooperation with you and your organization sincerely.
                                    </h4>

                                    <h4 style="margin: 0px; font-weight: lighter;">Best regards,</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">Dr. Cheng</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">Organizing Committee</h4>
                                    <h4 style="margin: 0px; font-weight: lighter;">AICS2019</h4>
                                </div>
                            </div>
                            </body>
                    """ % name
        return BODY_HTML

    # 计算机中文评审
    def Computer_chinese_evaluation(self, name):
        pass

    # 材料英文评审
    def Material_english_review(self, name):
        BODY_HTML = """<!DOCTYPE html>
                               <html lang="en">
                               <head>
                                   <meta charset="UTF-8">
                                   <title></title>
                               </head>
                               <body>
                               <div class="div_01" style="width: 600px; margin: 10px auto; font-family: Times New Roman; background-color: #cccccc;">
                                   <!--<p class="p_02" style="margin: 0px auto;"><img src="cid:image1" style="width: 600px; height:322px"></p>-->
                                   <h2 style="text-align: center; margin: 0px auto;">2019 International Conference on Material Engineering and Materials Science (MEMS2019)</h2>
                                   <h1 style="text-align: center; margin: 0px auto;">Invitation Letter</h1>
                                   <h2 style="text-align: center; margin: 0px auto;">for Scientific Committee Members and Guest Editors</h2>
                                   <div class="div_03" style="margin: 0px 5px; text-align: justify;">
                                       <h4 style="margin: 0px; font-weight: lighter;">Dear %s,</h4>

                                       <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                           &#8195;&#8195;It is my honor to contact you. This is Dr. Cheng from Organizing Committee of 2019 International Conference on Material Engineering and Materials Science (MEMS2019)!
                                        </h4>

                                       <h4 style="margin: 0px 0px 8px; font-weight: lighter;">
                                           &#8195;&#8195;In view of your achievements and contribution in academic researches, the influence on related research areas, MEMS2019 intends to invite you to join our Scientific Committee Members and Guest Editors Team, and act as our conferences academic guider promoting the development of the scientific fields jointly.
                                        </h4>

                                       <h4 style="margin: 0px; font-weight: lighter;">
                                           &#8195;&#8195;We are here sincerely inviting you to Scientific Committee Members and Guest Editors Team because of your great achievements and prestige in your areas.
                                        </h4>

                                       <h4 style="margin: 0px; font-weight: lighter;">
                                           &#8195;&#8195;If you feel interested in joining us, please fill in the application form below and send it to our mailbox <a href="mailto:easyace_cn@163.com">(easyace_cn@163.com)</a>. Your application will be handled in about 5 working days.
                                        </h4>

                                        <div style="width: 580px; height:680px; margin: 0px auto; border:#000000 solid 2px;">
                                            <div style="width: 578px; height:38px; margin: 0px auto; text-align:center; line-height:38px; font-weight:bold; border-bottom:1px solid;">Application Form</div>

                                            <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; line-height:30px; font-weight:bold; border-right:1px solid;">Name</div>
                                            <div style="width: 300px; height:31px; float:left; margin: 0px;"></div>
                                            <div style="width: 148px; height:30px; float:right; margin: 0px; border-left:1px solid;"></div>
                                            <div style="clear: both"></div>

                                            <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; font-weight:bold; line-height:30px; border-top:1px solid; border-right:1px solid; border-bottom:1px solid;">Email</div>
                                            <div style="width: 300px; height:31px; float:left; margin: 0px; border-top:1px solid; border-bottom:1px solid;"></div>
                                            <div style="width: 148px; height:30px; float:right; margin: 0px; border-left:1px solid;"></div>
                                            <div style="clear: both"></div>

                                            <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; font-weight:bold; line-height:30px; border-right:1px solid;">Title</div>
                                            <div style="width: 300px; height:31px; float:left; margin: 0px;"></div>
                                            <div style="width: 148px; height:30px; float:right; margin: 0px; border-left:1px solid;"></div>
                                            <div style="clear: both"></div>

                                            <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; font-weight:bold; line-height:30px; border-right:1px solid; border-top:1px solid; border-bottom:1px solid;">Institution</div>
                                            <div style="width: 300px; height:31px; float:left; margin: 0px; border-top:1px solid; border-bottom:1px solid;"></div>
                                            <div style="width: 148px; height:30px; float:right; margin: 0px; border-left:1px solid; border-bottom:1px solid;"></div>
                                            <div style="clear: both"></div>

                                            <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; font-weight:bold; line-height:30px; border-right:1px solid;">Research Fields</div>
                                            <div style="width: 449px; height:30px; float:left; margin: 0px;"></div>
                                            <div style="clear: both"></div>

                                            <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; font-weight:bold; line-height:30px; border-top:1px solid; border-bottom:1px solid; border-right:1px solid;">Address</div>
                                            <div style="width: 449px; height:30px; float:left; margin: 0px; border-top:1px solid; border-bottom:1px solid;"></div>
                                            <div style="clear: both"></div>

                                            <div style="width: 128px; height:30px; float:left; margin: 0px; text-align:center; font-weight:bold; line-height:30px; border-right:1px solid;">Choose to be</div>
                                            <div style="width: 449px; height:30px; float:left; margin: 0px; text-align:center; line-height:30px;">Scientific Committee Member □&#8195;&#8195;Guest Editor □&#8195;&#8195;Either □</div>
                                            <div style="clear: both"></div>

                                            <div style="width: 578px; height:38px; margin: 0px auto; text-align:center; font-weight:bold; line-height:38px; border-top:1px solid; border-bottom:1px solid;">Publication List</div>
                                            <div style="width: 578px; height:100px; margin: 0px auto; text-align:center;"></div>
                                            <div style="width: 578px; height:38px; margin: 0px auto; text-align:center; font-weight:bold; line-height:38px; border-top:1px solid; border-bottom:1px solid;">Achievements in Conferences and Journals </div>
                                            <div style="width: 578px; height:100px; margin: 0px auto; text-align:center;"></div>
                                            <div style="width: 578px; height:38px; margin: 0px auto; text-align:center; font-weight:bold; line-height:38px; border-top:1px solid; border-bottom:1px solid;">Other</div>
                                            <div style="width: 578px; height:100px; margin: 0px auto; text-align:center;"></div>
                                        </div>

                                       <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                           &#8195;&#8195;And we are looking for international cooperations with institutes, colleges, agencies in related areas and holding the conference together; if you feel it is the chance for us to work closely in academic communication, please do recommend MEMS2019 to your organization.
                                        </h4>

                                       <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                           &#8195;&#8195;If you have any problems or suggestions on the cooperation, please feel free to contact me!
                                        </h4>

                                       <h4 style="margin: 8px 0px 8px; font-weight: lighter;">
                                           &#8195;&#8195;MEMS2019 welcomes the cooperation with you and your organization sincerely.
                                       </h4>

                                       <h4 style="margin: 0px; font-weight: lighter;">Best regards,</h4>
                                       <h4 style="margin: 0px; font-weight: lighter;">Dr. Cheng</h4>
                                       <h4 style="margin: 0px; font-weight: lighter;">Organizing Committee</h4>
                                       <h4 style="margin: 0px; font-weight: lighter;">MEMS2019</h4>
                                   </div>
                               </div>
                               </body>
                       """ % name
        return BODY_HTML

    # 材料中文评审
    def Material_chinese_review(self, name):
        pass

    def send_email(self, row_list):
        theme_str = self.theme[random.randint(0, len(self.theme) - 1)]
        meil_str = self.meil_list[random.randint(0, len(self.meil_list) - 1)]

        SENDER = meil_str  # 发送邮箱的用户名
        SENDERNAME = 'Organizing Committee'  # 发件人姓名
        RECIPIENT = [row_list[1], ]  # 发送到的邮件，可是list, 但注意时msg['To']值必须是字符串用,隔开

        USERNAME_SMTP = "AKIAIR45WGMMJV5EFRYENA"  # 带有邮件权限的 IAM 帐号
        PASSWORD_SMTP = "Ai7Rfn5LSMEPi/8EJivJr4PxniSltb65Wo/RnfIEhsHb"  # 带有邮件权限的 IAM 密码
        HOST = "email-smtp.us-east-1.amazonaws.com"  # Amazon SES SMTP 终端节点
        PORT = 25  # SMTP 客户端连接到STARTTLS 端口 25、587 或 2587 上的 Amazon SES SMTP 终端节点
        SUBJECT = theme_str  # 主题

        # 测试邮件正文
        BODY_TEXT = ("Amazon SES Test\r\n"
                     "This email was sent through the Amazon SES SMTP ")

        msg = MIMEMultipart('alternative')
        msg['Subject'] = SUBJECT
        msg['From'] = email.utils.formataddr((SENDERNAME, SENDER))  # 发件人邮箱
        # msg['To'] = RECIPIENT  # 收件人邮箱
        msg['To'] = ','.join(RECIPIENT)  # 收件人邮箱,不同收件人邮箱之间用,分割,再组合为字符串

        # 邮件内容，记得一定要用plain传入
        # part1 = MIMEText(BODY_TEXT, 'plain')

        if self.cov_str == "土木英文":
            part2 = MIMEText(self.civil_english_email(row_list[0]), 'html', "utf-8")
        elif self.cov_str == "土木中文":
            part2 = MIMEText(self.civil_chinese_email(row_list[0]), 'html', "utf-8")
        elif self.cov_str == "地质英文":
            part2 = MIMEText(self.geology_english_email(row_list[0]), 'html', "utf-8")
        elif self.cov_str == "地质中文":
            part2 = MIMEText(self.geology_chinese_email(row_list[0]), 'html', "utf-8")
        elif self.cov_str == "计算机英文评审":
            part2 = MIMEText(self.Computer_english_evaluation(row_list[0]), 'html', "utf-8")
        elif self.cov_str == "计算机中文评审":
            part2 = MIMEText(self.Computer_chinese_evaluation(row_list[0]), 'html', "utf-8")
        elif self.cov_str == "材料英文评审":
            part2 = MIMEText(self.Material_english_review(row_list[0]), 'html', "utf-8")
        elif self.cov_str == "材料中文评审":
            part2 = MIMEText(self.Material_chinese_review(row_list[0]), 'html', "utf-8")
        else:

            print("请确认文件名: {}".format(self.csv_name_list))
            return
        # msg.attach(part1)
        msg.attach(part2)

        # 邮件带图片时使用二进制
        if "土木" in self.cov_str:
            fp = open(r'./building.jpg', 'rb')
        elif "地质" in self.cov_str:
            fp = open(r'./geological.jpg', 'rb')
        elif self.cov_str in "计算机英文评审" or self.cov_str in "计算机中文评审":
            pass
        elif self.cov_str in "材料英文评审" or self.cov_str in "材料中文评审":
            pass
        else:
            print("无相应的图片")
            return

        if "土木" in self.cov_str or "地质" in self.cov_str:
            msgImage = MIMEImage(fp.read())
            fp.close()

            # 定义图片 ID，在 HTML 文本中引用
            msgImage.add_header('Content-ID', '<image1>')
            msg.attach(msgImage)

        try:
            server = smtplib.SMTP(HOST, PORT)
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(USERNAME_SMTP, PASSWORD_SMTP)

            # 发送的时候需要将收件人和抄送者全部添加到函数第二个参数里
            server.sendmail(SENDER, RECIPIENT, msg.as_string())
            server.close()
        except Exception as e:
            try:
                if e == "[SSL: UNKNOWN_PROTOCOL] unknown protocol (_ssl.c:645)":
                    server = smtplib.SMTP(HOST, 465)
                    server.ehlo()
                    server.starttls()
                    server.ehlo()
                    server.login(USERNAME_SMTP, PASSWORD_SMTP)

                    # 发送的时候需要将收件人和抄送者全部添加到函数第二个参数里
                    server.sendmail(SENDER, RECIPIENT, msg.as_string())
                    server.close()
                else:
                    server = smtplib.SMTP(HOST, PORT)
                    server.starttls()
                    server.login(USERNAME_SMTP, PASSWORD_SMTP)

                    # 发送的时候需要将收件人和抄送者全部添加到函数第二个参数里
                    server.sendmail(SENDER, RECIPIENT, msg.as_string())
                    server.close()
            except Exception as e:
                print("Error: ", e)
                row_list[1] = row_list[1] + ':{}'.format(str(e))
                self.ctrl_function(self.cov_str + "发送失败", row_list)
            else:
                print("Email 成功!")
                self.ctrl_function(self.cov_str + "发送成功", row_list)
        else:
            print("Email 成功!")
            self.ctrl_function(self.cov_str + "发送成功", row_list)

    def main_sendmail(self, old_queue):
        while True:
            # 子进程等待任务加入old_queue队列中, 以便获取
            proxy = old_queue.get()
            if proxy == 0: break
            time.sleep(self.dwell_time)
            self.send_email(proxy)

    def main(self):
        # 需处理的邮件
        old_queue = Queue()

        # 进程池列表
        works = []
        for _ in range(self.int_initnum):
            # 循环 Process生成5个子进程, 加入进程池
            works.append(Process(target=self.main_sendmail, args=(old_queue,)))

        for work in works:
            # 循环开启子进程
            work.start()

        for proxy in self.proxies:
            # 将任务加入old_queue队列
            old_queue.put(proxy)

        for work in works:
            old_queue.put(0)

        for work in works:
            # 等待子进程执行完后关闭
            work.join()


if __name__ == "__main__":
    # 停顿时间
    dwell_time = 5

    # 子进程数量
    int_initnum = 5

    # 异常邮箱是否通过
    # False 或者数字:0 在不确认邮箱格式是否正确请使用;
    # True 或者数字:1 在确认全部正确情况下使用;
    Abnormal_through = False

    # 发件人列表
    meil_list = ["gteew@easyaca.cn', 'ewffdfs@easyaca.cn', 'rgrtfvcx@easyaca.cn"]

    # 文本模板文件名
    csv_name_list = ["土木英文", "土木中文", "地质英文", "地质英文", "计算机英文评审", "计算机中文评审", "材料英文评审", "材料中文评审"]

    sender = SendEmail(dwell_time, int_initnum, Abnormal_through, csv_name_list, meil_list)
    sender.main()
