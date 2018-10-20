#导入模块和打开表单
import xlwt,xlrd
import datetime

wb=xlrd.open_workbook('blsx.xls')
sheet=wb.sheet_by_index(0)

f=open('blsx.txt','r+')

#写病程时间
ryrq=xlrd.xldate_as_tuple(sheet.cell_value(5,6),wb.datemode)
ryrq=datetime.date(*ryrq[:3])

cyrq=sheet.cell_value(6,6)
if cyrq=='':  
    cyrq=datetime.datetime.now()
    cyrq=cyrq.date()
    zyzt=1 #回车标记，在院
else:  #有输入日期
    cyrq=xlrd.xldate_as_tuple(sheet.cell_value(6,6),wb.datemode)
    cyrq=datetime.date(*cyrq[:3])
    zyzt=0
    zyts=(cyrq-ryrq).days #住院天数
print('出入院日期，都是date：',ryrq,cyrq)

ns=str(int(sheet.cell_value(7,6))) #年数 *
zs=int(sheet.cell_value(8,6)) #注射 *
##ryyy=sheet.cell_value(9,6)##入院用药
##ryjc=sheet.cell_value(10,6) ##入院检查10
##zzcf=sheet.cell_value(11,6) ##主治11
##zrcf=sheet.cell_value(12,6) ##主任12
bcqm='			                                              医师签名：庄松源'
qbzl=''

if int(zs)==0:
    zswb='四肢未见注射疤痕。'
else:
    zswb='四肢可见多处注射疤痕，无红肿。'

def zl(mb,lb):  #返回各种措施的组合。mb从0开始，定位黄色区域内容。lb从1-3，定位白色区域
    global qbzl
    zlwb=''
    xz=sheet.cell(mb+9,6).value #现有治疗，注意此值为空时
    xztype=sheet.cell(mb+9,6).ctype
    if xztype==2: #是一个数，即是只有一个选择
        xz=str(int(xz))
    if not xz=='':
        xzl=xz.split('+')
        xzlint=[]
        for i in xzl:
            xzlint.append(int(i))
        for i in xzlint:  #迭代选择的每一个数字
            h=int(i)  #行
            zlwb=zlwb+sheet.cell(h,lb).value+'，' #合并所有选择的文本，选择类别
    elif lb!=2:        
        zlwb='患者病情平稳，继续美沙酮替代递减治疗，加强心理辅导及行为训练。'
    fc=sheet.cell_value(mb+9,7)
    zlcs=fc+zlwb[:-1]+'。' #治疗措施
    if (xz!='' and lb!=2): #汇集总的治疗
        qbzl=qbzl+zlwb
    print(zlcs)
    return zlcs

#fc='辅助检查结果回报：患者拒查心电图，三大常规、血液生化检查均无明显异常。'


bc1=ryrq+datetime.timedelta(days=1) #入院第2日
bc2=datetime.datetime.strftime(bc1,'%Y-%m-%d 10:30') #转文本
print(bc2,'            李赛民主治医师查房记录')
bct='%s            李赛民主治医师查房记录' % bc2  #标题
if bc1<cyrq or (bc1==cyrq and zyzt==1): #未出院，或为当天写记录
    ryzl=zl(0,1) #0+9即是第10行
    zyzl=zl(2,3)
    if sheet.cell(11,6).ctype!=0:
        dj=''
    else:
        dj='逐渐递减美沙酮用量，'
        zyzl=''
    #ns,ryzl,zswb,dj,zyzl
    bcb="    患者因“反复滥用海洛因%s年”入院。入院后给予%s今随李赛民主治医师查看患者，患者精神及体力好，未诉打哈欠、流泪等戒断不适，食纳可，睡眠一般，大便干结，小便正常。查体：生命体征平稳、神志清楚，面色晦暗，双侧瞳孔等大等圆，直径约3.0mm，对光反射灵敏，心肺听诊无明显异常，腹平、软，无压痛。%s未引出幻觉及妄想，自知力及定向力正常。李赛民主治医师查看病人后指示：同意目前诊断和治疗，%s%s并给予脑功能理疗调节神经功能。" % (ns,ryzl,zswb,dj,zyzl)
    f.write(bct+'\n')
    f.write(bcb+'\n')
    f.write(bcqm+'\n')


bc1=ryrq+datetime.timedelta(days=2) #入院第3日
bc2=datetime.datetime.strftime(bc1,'%Y-%m-%d 10:30')
print(bc2,'            段爱明主任医师查房记录')
bct='%s            段爱明主任医师查房记录' % bc2
if bc1<cyrq or (bc1==cyrq and zyzt==1):
    ryjc='辅助检查：'+zl(1,2)
    zyzl=zl(3,3)
    #zswb,ryjc,zyzl
    bcb="     今日随段爱明副主任医师查看病人，患者未诉打哈、流泪等戒断不适，精神好，食纳可，睡眠一般，大便干结，小便正常。查体：生命体征平稳、神志清楚，面色晦暗，双侧瞳孔等大等圆，直径约3.0mm，对光反射灵敏，心肺听诊无明显异常，腹平、软，无压痛。%s未引出幻觉及妄想，自知力及定向力正常。%s段爱明副主任医师查看病人后指示：同意入院诊断和目前诊疗措施；%s加强心理辅导，增强患者戒毒的信心和决心。同时叮嘱患者严守院规，禁止串房和聚众闹事；以上指示均照执行。" % (zswb,ryjc,zyzl)
    f.write(bct+'\n')
    f.write(bcb+'\n')
    f.write(bcqm+'\n')



i=5 #天数
c=4 #第一次日常查房的mb数值
bc1=ryrq+datetime.timedelta(days=i)
while bc1<cyrq or (bc1==cyrq and zyzt==1): #日期相同，在院
    bc2=datetime.datetime.strftime(bc1,'%Y-%m-%d 10:30')    
    ys=((bc1-ryrq).days+1)%30
    if ys<3:
        print('余数是',ys)
        bct='%s            阶段小结' % bc2
        bcb=''
    else:
        bct=bc2
        print(bct)  
        zyzl=zl(c,3)
        bcb="    今日查房，患者未诉打哈欠、流泪等戒断不适，精神好，食纳可，睡眠一般，大小便正常。查体：生命体征平稳、神志清楚，面色晦暗，双侧瞳孔等大等圆，直径约3.0mm，对光反射灵敏，心肺腹部检查未见明显异常。未引出幻觉及妄想，自知力及定向力正常。%s" % zyzl
        
    f.write(bct+'\n')
    f.write(bcb+'\n')
    f.write(bcqm+'\n')           
    i=i+3
    c=c+1
    bc1=ryrq+datetime.timedelta(days=i)

if bc1>=cyrq and zyzt==0:  #已出院，未次病程时间为cyrq
    bc2=datetime.datetime.strftime(cyrq,'%Y-%m-%d 20:30')
    bct=bc2
    print(bct)
    bcb1='    今日查房，'
    zyzl=zl(c,3)
    #zyzl
    bcb2='患者精神好，未诉明显不适，食纳可，睡眠一般，大小便正常。查体：生命体征平稳，神志清楚，双侧瞳孔等大等圆，直径约3.0mm，对光反射灵敏。心肺听诊未见异常，腹平、软，无压痛。未引出幻觉或妄想，自知力及定向力正常，意志力减退。%s' % zyzl
    bcb=bcb1+bcb2
    cy=int(sheet.cell_value(6,7))
    if cy==4:
        bcb3='患者因个人事务主动要求于今日出院，经劝说无效，予以办理。'
        cyfs='因个人事务而主动要求出院。'
    else:
        bcb3='患者病情好转，予以办理出院。'
        cyfs='好转出院。'        
    bcb=bcb1+bcb2+bcb3    
    f.write(bct+'\n')
    f.write(bcb+'\n')
    f.write(bcqm+'\n')

    #ns,zswb,ryjc,qbzl,zyts,cyfs
    ryqk='患者因“反复滥用海洛因%s年”入院，中断烫吸海洛因6-8小时后出现戒断症状，吸食后缓解。入院查体：神志清楚，慢性病面容，双瞳孔等大等圆，直径3.0mm，对光反射灵敏。心肺听诊未见异常，腹平、软，无压痛。%s未引出幻觉或妄想，自知力及定向力正常，意志力减退。尿吗啡检测：阳性。患者入院后%s给予%s并辅以心理治疗及理疗。患者住院%d天，%s' % (ns,zswb,ryjc,qbzl,zyts,cyfs)
    cyqk=bcb2
    f.write('\n')
    f.write(ryqk+'\n')
    f.write('\n')
    f.write(cyqk+'\n')

f.close()
  


