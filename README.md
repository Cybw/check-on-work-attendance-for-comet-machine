# check-on-work-attendance-for-comet-machine
A little script for count lab student attendance ,a job for superviser. Data is output from a Comet(科密) machine per week. With slow algorithm,but work.

Caution!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
The demo.xls is encoding by 'gbk',but the truely data I used is 'cp1252'. So , if u want to run the demo ... 
please use :
data = xlrd.open_workbook(filename + '.xls', encoding_override='gbk')
and
with codecs.open(filename + '.csv', 'w', encoding='gbk') as f:
Thanks!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

使用说明：
    1、“coursetab.csv”文件为教务处公布的每日上课时间段；
    2、“coursetime.csv”文件为参与打卡的每位同学的上课时间（所有
    参与打卡的同学必须都包含在表内，学校助管、公共活动等都按上课处理）
	表中从周一到周五，‘o’为不上课，‘1’为上课；
    3、使用前请将考勤报表上传至脚本同一目录下；
    4、将文件名替换为需要统计的表格名（不带.xls），确认是否需要包括
    上课时间，点击运行即可。
    
