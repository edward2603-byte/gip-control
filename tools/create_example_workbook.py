from datetime import date
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).with_name('vendor')))
import xlsxwriter
out=Path(__file__).parents[1]/'examples'/'GIP_Navigator_STAGE0_DEMO.xlsx'
if len(sys.argv)==3: out=Path(sys.argv[2])
out.parent.mkdir(parents=True,exist_ok=True); w=xlsxwriter.Workbook(out); w.set_calc_mode('auto')
h=w.add_format({'bold':True,'font_color':'white','bg_color':'#C0504D'}); d=w.add_format({'num_format':'dd.mm.yyyy'}); y=w.add_format({'num_format':'dd.mm.yyyy','bg_color':'#FFF2CC'}); yi=w.add_format({'bg_color':'#FFF2CC'})
home=w.add_worksheet('00_Главная');kp=w.add_worksheet('01_КП');sr=w.add_worksheet('03_СР_проекта');g=w.add_worksheet('05_График');c=w.add_worksheet('06_Контроль')
home.set_column('A:A',105);home.write('A1','GIP Navigator — этап 0');home.write('A3','Жёлтые ячейки на листе 03_СР_проекта редактируются. Прогноз, контроль и диаграмма пересчитываются без VBA.')
kp.write_row('A1',['Этап','Наименование','Начало КП','Окончание КП','Стоимость','Прогноз','Статус'],h)
for r,x in enumerate([['КП-1','Предпроектные проработки',date(2026,2,2),date(2026,2,14),1200000],['КП-2','Техническое проектирование',date(2026,2,15),date(2026,2,28),1800000]],1):
 kp.write(r,0,x[0]);kp.write(r,1,x[1]);kp.write_datetime(r,2,x[2],d);kp.write_datetime(r,3,x[3],d);kp.write_number(r,4,x[4])
kp.write_formula('F2',"=MAX('03_СР_проекта'!$O$2:$O$5)",d);kp.write_formula('F3',"=MAX('03_СР_проекта'!$O$6:$O$7)",d);kp.write_formula('G2','=IF(F2>D2,"Критично: позднее КП","Норма")');kp.write_formula('G3','=IF(F3>D3,"Критично: позднее КП","Норма")');kp.set_column('A:G',18)
hdr=['Код','Работа','Этап','Предшественник','Связь','Lag','План','Длит.','Факт. начало','Факт. окончание','Ожид. окончание','Состояние ИД','Ожид. дата ИД','Прогноз начала','Прогноз окончания','Статус'];sr.write_row('A1',hdr,h)
rows=[['SR-01','Получение ИД','КП-1','','FS',0,date(2026,2,2),3,'Получены',None],['SR-02','Инженерные изыскания','КП-1','','FS',0,date(2026,2,2),5,'Получены',None],['SR-03','Гидравлический расчёт','КП-1','SR-01','FS',0,date(2026,2,5),4,'Ожидание',date(2026,2,7)],['SR-04','Предварительные решения','КП-1','SR-03','FS',0,date(2026,2,9),3,'Получены',None],['SR-05','Техническое проектирование','КП-2','SR-04','FS',0,date(2026,2,15),6,'Получены',None],['SR-06','Материалы на согласование','КП-2','SR-05','SS',3,date(2026,2,18),4,'Получены',None]]
for i,x in enumerate(rows,1):
 r=i+1;sr.write_row(i,0,x[:6]);sr.write_datetime(i,6,x[6],y);sr.write_number(i,7,x[7],yi)
 for q in (8,9,10):sr.write_blank(i,q,None,y)
 sr.write(i,11,x[8],yi);(sr.write_datetime(i,12,x[9],y) if x[9] else sr.write_blank(i,12,None,y))
 f=f'=IF(I{r}<>"",I{r},IF(L{r}="Нет данных","",IF(D{r}="",IF(L{r}="Ожидание",M{r},G{r}),IF(INDEX($O$2:$O$7,MATCH(D{r},$A$2:$A$7,0))="","",MAX(G{r},IF(E{r}="FS",INDEX($O$2:$O$7,MATCH(D{r},$A$2:$A$7,0))+1+F{r},INDEX($N$2:$N$7,MATCH(D{r},$A$2:$A$7,0))+F{r}),IF(L{r}="Ожидание",M{r},0))))))'
 sr.write_formula(i,13,f,d);sr.write_formula(i,14,f'=IF(N{r}="","",IF(J{r}<>"",J{r},IF(K{r}<>"",K{r},N{r}+H{r}-1)))',d);sr.write_formula(i,15,f'=IF(L{r}="Нет данных","Неопределённость: нет даты ИД",IF(O{r}="","Невозможно прогнозировать",IF(O{r}>VLOOKUP(C{r},\'01_КП\'!$A$2:$D$3,4,FALSE),"Критично: позднее КП","Норма")))')
sr.data_validation('L2:L7',{'validate':'list','source':['Получены','Ожидание','Нет данных']});sr.conditional_format('O2:O7',{'type':'3_color_scale'});sr.set_column('A:A',10);sr.set_column('B:B',30);sr.set_column('C:O',15);sr.set_column('P:P',32)
c.write_row('A1',['Код','Работа','Прогноз','Сигнал'],h)
for r in range(2,8):
 for i,col in enumerate('ABOP'):c.write_formula(r-1,i,f"='03_СР_проекта'!{col}{r}",d if col=='O' else None)
c.set_column('A:D',28)
g.write_row('A1',['Тип','Наименование','Начало','Длительность','КП — срок','СР — срок'],h)
for i,r in enumerate([2,3],1):
 g.write(i,0,'КП');g.write_formula(i,1,f"='01_КП'!B{r}");g.write_formula(i,2,f"='01_КП'!C{r}",d);g.write_formula(i,3,f"='01_КП'!D{r}-'01_КП'!C{r}+1")
for i,r in enumerate(range(2,8),3):
 g.write(i,0,'СР');g.write_formula(i,1,f"='03_СР_проекта'!B{r}");g.write_formula(i,2,f"='03_СР_проекта'!N{r}",d);g.write_formula(i,3,f"='03_СР_проекта'!O{r}-'03_СР_проекта'!N{r}+1")
for r in range(2,10):g.write_formula(r-1,4,f'=IF(A{r}="КП",D{r},NA())');g.write_formula(r-1,5,f'=IF(A{r}="СР",D{r},NA())')
ch=w.add_chart({'type':'bar','subtype':'stacked'});ch.add_series({'categories':"='05_График'!$B$2:$B$9",'values':"='05_График'!$C$2:$C$9",'fill':{'none':True},'border':{'none':True}});ch.add_series({'name':'КП — договорные сроки','categories':"='05_График'!$B$2:$B$9",'values':"='05_График'!$E$2:$E$9",'fill':{'color':'#7B3F98'},'border':{'color':'#125AB6'}});ch.add_series({'name':'СР — рабочие сроки','categories':"='05_График'!$B$2:$B$9",'values':"='05_График'!$F$2:$F$9",'fill':{'color':'#F5A623'},'border':{'color':'#7B3F98'}});ch.set_title({'name':'Диаграмма Ганта: этапы КП и СР'});ch.set_legend({'position':'bottom'});ch.set_x_axis({'min':46045,'max':46110,'major_unit':10,'num_format':'dd.mm'});ch.set_y_axis({'reverse':True});g.insert_chart('A12',ch,{'x_scale':1.55,'y_scale':1.4});g.set_column('A:F',20)
w.close();print('Создано:',out.resolve())
