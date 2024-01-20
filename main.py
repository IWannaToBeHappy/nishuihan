from pandas import DataFrame,read_excel,isnull
from bisect import bisect,insort
from re import sub
import itertools
import copy
from tqdm import tqdm

source =   read_excel('./基础设置.xlsx','候选内功')
attacker_panel = read_excel('./基础设置.xlsx','无内功面板')
target_panel = read_excel('./基础设置.xlsx','目标数值')
周天 = read_excel('./基础设置.xlsx','周天加成')
种类 = read_excel('./基础设置.xlsx','内功种类加成')
args = read_excel('./基础设置.xlsx','计算公式') 

内功种类加成 = {}
for i in range(len(种类)):
    内功种类加成[种类.iloc[i]['内功']] = (种类.iloc[i]['内功加成'],种类.iloc[i]['灵韵加成'])


### 用户交互1----------------------------------------------------------------------------------------------
print('Tips:\n \
      必做：请根据自身情况先修改 基础设置.xlsx 中的候选内功、无内功面板\n \
      进阶设置（选做）：请在 内功种类加成表 中手动设置断金戈、冲虚、点明炉、灭阵四个内功的内功加成，方法是8*对应类型伤害占比，举例：若无内功情况下爆发伤害占比为33%，则断金戈一栏可修改为12*33%=4,灵韵栏减半，修改为2\n \
      版本更新进阶(选做):可根据主播、大神实测数据修改目标数值、内功种类加成、周天加成表中内容\n \
      强烈不建议：修改计算公式表中内容\n')
try:
    num_内功 = int(input('你可以装下几个内功？(请输入5 or 6)\n'))
except:
    print("无效输入，默认使用5内功组合版本\n")
    num_内功 = 5


try:
    职业 = int(input('请选择你的职业：碎梦请输1 血河请输2 神相请输3 龙吟请输4 九灵请输5 铁衣请输6 素问请输7\n'))
except:
    print('无效输入 默认为九灵 :p\n')
    职业=5
try:
    偷师 = int(input('请选择你的偷师大特质（仅影响大特质）：碎梦请输1 血河请输2 神相请输3 龙吟请输4 九灵请输5 铁衣请输6 素问请输7\n'))
except:
    print('无效输入 默认选龙吟\n')
    偷师=4

### -----------------------------------------------------------------------------------------------
    
class 内功():
    def __init__(self,iloc) -> None:
        self.种类 = iloc['名称']
        self.金 = 0 if isnull(iloc['金']) else iloc['金']
        self.木 = 0 if isnull(iloc['木']) else iloc['木']
        self.火 = 0 if isnull(iloc['火']) else iloc['火']
        self.灵韵 = False
        self.词条 = {'破防':0,
                   '力量':0,
                   '气海':0,
                   '命中':0,
                   '攻击':0,
                   '根骨':0,
                   '身法':0,
                   '会心':0,
                   '最大攻击':0,
                   '最小攻击':0,
                   '全元素攻击':0,
                   '单元素攻击':0,
                   '首领克制':0,
                   '耐力':0
                   }
        for k,v in iloc.items():
            if v=='灵韵':
                self.灵韵 = True
                continue
            if '词条' not in k or isnull(v):continue
            self.词条[v]+=iloc[sub('词条','数值',k)]
        self.种类加成 = 内功种类加成[self.种类][0]
        if self.灵韵:
            self.种类加成 += 内功种类加成[self.种类][1]
        self.panel = self.cal_面板加成()
    
    def cal_面板加成(self):
        global 职业,偷师
        tmp_panel = Panel(攻击=0,命中=0,会心=0,元素攻击=0,克制_数值=0,破防=0)
        tmp_panel.破防+=self.词条['破防']
        tmp_panel.命中+=self.词条['命中']
        tmp_panel.攻击+=self.词条['攻击']
        tmp_panel.会心+=self.词条['会心']
        tmp_panel.攻击+=(self.词条['最大攻击']+self.词条['最小攻击'])/2
        tmp_panel.元素攻击+=self.词条['全元素攻击']+self.词条['单元素攻击']
        tmp_panel.克制_数值+=self.词条['首领克制']

        # 气海、力量、根骨、耐力、身法
        if 职业 in (1,2,4,5):
            tmp_panel.破防+=2*(self.词条['力量'])
            tmp_panel.攻击+=4.8*(self.词条['力量'])
        else:
            tmp_panel.破防+=2*(self.词条['气海'])
            tmp_panel.攻击+=4.8*(self.词条['气海'])
        tmp_panel.命中+=0.83*self.词条['身法']
        tmp_panel.会心+=1.69*self.词条['身法']
        if 1 in (职业,偷师):
            tmp_panel.会心+=0.6*self.词条['身法']
        if 2 in (职业,偷师):
            tmp_panel.攻击+=1.65*self.词条['根骨']
        if 3 in (职业,偷师):
            tmp_panel.命中+=0.83*(self.词条['气海']+self.词条['力量'])
        if 4 in (职业,偷师):
            tmp_panel.破防+=2*self.词条['根骨']
        if 5 in (职业,偷师):
            tmp_panel.攻击+=1.65*self.词条['耐力']
        return tmp_panel
        

    def __str__(self) -> str:
        s = self.种类+':\t'+str(self.金)+'金 '+str(self.木)+'木 '+str(self.火)+'火 '
        if self.灵韵:
            s+='灵韵\t'
        for k,v in self.词条.items():
            if v==0 or v==None:continue
            s+=k+':'+str(v)+'\t'
        return s
    
class Panel():
    def __init__(self,攻击=None,命中=None,会心=None,会伤=None,破盾=None,元素攻击=None,克制_数值=None,克制_百分比=None,破防=None,单体占比=None,群体占比=None,爆发占比=None,持续占比=None,
        防御=None,气盾=None,格挡=None,会心抗性=None,会心防御=None,元素抗性=None,抵御=None,伤害减免=None,单体减免=None,群体减免=None,爆发减免=None,持续减免=None
    ) -> None:
        self.攻击=攻击
        self.命中=命中
        self.会心=会心
        self.会伤=会伤
        self.破盾=破盾
        self.元素攻击=元素攻击
        self.克制_数值=克制_数值
        self.克制_百分比=克制_百分比
        self.破防=破防
        self.单体占比=单体占比
        self.群体占比=群体占比
        self.爆发占比=爆发占比
        self.持续占比=持续占比

        self.防御=防御
        self.气盾=气盾
        self.格挡=格挡
        self.会心抗性=会心抗性
        self.会心防御=会心防御
        self.元素抗性=元素抗性
        self.抵御=抵御
        self.伤害减免=伤害减免
        self.单体减免=单体减免
        self.群体减免=群体减免
        self.爆发减免=爆发减免
        self.持续减免=持续减免
    
    def copy_add(self,panels):
        tmp_panel = copy.copy(self)
        for attr in ('攻击','命中','会心','元素攻击','克制_数值','破防'):
            value = getattr(tmp_panel,attr)
            for panel in panels:
                value+=getattr(panel,attr)
            setattr(tmp_panel,attr,value)
        return tmp_panel


def cal_damage(attacker,defender):
    会心率=(args.iloc[4]['参数1']*(attacker.会心-defender.会心抗性)+args.iloc[4]['参数2'])/(attacker.会心-defender.会心抗性+args.iloc[4]['参数3'])/100
    命中率=min(1,0.95+(args.iloc[3]['参数1']*attacker.命中/(attacker.命中+args.iloc[3]['参数2'])-args.iloc[3]['参数1']*defender.格挡/(defender.格挡+args.iloc[3]['参数2']))/100)
    抗性减免=defender.元素抗性/(defender.元素抗性+args.iloc[2]['参数1'])
    防御减免=(defender.防御-attacker.破防)/(defender.防御+args.iloc[1]['参数1'])
    伤害=((args.iloc[0]['参数1']+(attacker.攻击+min(attacker.破盾-defender.气盾,0)+attacker.克制_数值-defender.抵御)*(1-防御减免)+attacker.元素攻击*(1-抗性减免)))*((attacker.会伤-defender.会心防御)*会心率/100+1)*命中率
    return 伤害

class Scorer():
    def __init__(self) -> None:
        周天加成 = {}
        for i in range(len(周天)):
            周天加成[周天.iloc[i]['周天']]=周天.iloc[i]['百分比加成']
        self.周天加成 = 周天加成
        try:
            print('请输入想要看到的内功组合数量:')
            self.top = int(input())
        except:
            self.top = 1

    def get_zhoutian_score(self,内功集):
        百分比加成 = 0
        jin = sum([neigong.金 for neigong in 内功集])//4
        mu = sum([neigong.木 for neigong in 内功集])//4
        huo = sum([neigong.火 for neigong in 内功集])//4
        if jin==1:百分比加成+=self.周天加成['1金']
        if jin==2:百分比加成+=self.周天加成['2金']
        if jin>=3:百分比加成+=self.周天加成['3金']
        if mu==1:百分比加成+=self.周天加成['1木']
        if mu==2:百分比加成+=self.周天加成['2木']
        if mu>=3:百分比加成+=self.周天加成['3木']
        if huo==1:百分比加成+=self.周天加成['1火']
        if huo==2:百分比加成+=self.周天加成['2火']
        if huo>=3:百分比加成+=self.周天加成['3火']
        return 百分比加成

    def get_best_combination(self,攻击者面板,目标面板,候选内功集,num_内功):
        top_damage = []
        top_组合 = []
        for 内功组合 in tqdm(itertools.combinations(候选内功集,num_内功)):
            if len(set([内功.种类 for 内功 in 内功组合]))!=num_内功:continue
            # 计算攻击者面板
            组合面板 = 攻击者面板.copy_add([内功.panel for 内功 in 内功组合])
            damage = cal_damage(组合面板,目标面板)
            # 内功种类加成以及周天加成
            damage *= (1+self.get_zhoutian_score(内功组合)/100)
            for i in 内功组合:
                damage *= (1+i.种类加成/100)
                        
            position = bisect(top_damage,damage)
            if position==0 and len(top_damage)>self.top:continue
            top_damage.insert(position,damage)
            top_组合.insert(position,内功组合)
            if len(top_damage)>self.top:
                top_damage.pop(0)
                top_组合.pop(0)
        for i in range(-1,-self.top-1,-1):
            print(str(top_damage[i]))
            for 内功 in top_组合[i]:
                print(str(内功))
            print('\n')



try:
    候选内功集 = []
    for i in range(len(source)):
        候选内功集.append(内功(source.iloc[i]))

    攻击者属性 = {}
    for i in range(len(attacker_panel)):
        攻击者属性[attacker_panel.iloc[i]['属性']] = attacker_panel.iloc[i]['数值']
    攻击者面板 = Panel(
        攻击=攻击者属性['攻击'],
        命中=攻击者属性['命中'],
        会心=攻击者属性['会心'],
        会伤=攻击者属性['会心伤害'],
        破盾=攻击者属性['破盾'],
        元素攻击=攻击者属性['元素攻击'],
        克制_数值=攻击者属性['首领克制/数值'],
        克制_百分比=攻击者属性['首领克制/百分比'],
        破防=攻击者属性['破防'],
        单体占比=攻击者属性['单体占比'],
        群体占比=攻击者属性['群体占比'],
        爆发占比=攻击者属性['爆发占比'],
        持续占比=攻击者属性['持续占比'],
    )

    目标属性 = {}
    for i in range(len(target_panel)):
        目标属性[target_panel.iloc[i]['属性']] = target_panel.iloc[i]['数值']
    目标面板 = Panel(
        防御=目标属性['防御'],
        气盾=目标属性['气盾'],
        格挡=目标属性['格挡'],
        会心抗性=目标属性['会心抗性'],
        会心防御=目标属性['会心防御'],
        元素抗性=目标属性['元素抗性'],
        抵御=目标属性['抵御'],
        伤害减免=目标属性['伤害减免'],
        单体减免=目标属性['单体减免'],
        群体减免=目标属性['群体减免'],
        爆发减免=目标属性['爆发减免'],
        持续减免=目标属性['持续减免'],
    )

    Scorer().get_best_combination(攻击者面板,目标面板,候选内功集,num_内功)
    print('回车退出程序')
    input()
except Exception as e:
    input(e)