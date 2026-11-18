"""Build a concrete example workbook from a validated КП source.

This helper is deliberately kept out of the production import workflow.  It
generates examples only and must not be used to update a GIP's working book.
"""
from __future__ import annotations
import argparse, sys
from datetime import date
from pathlib import Path
sys.path.insert(0, str(Path(__file__).with_name('vendor')))
import xlsxwriter
from gipnav_import import preview

ROOT=Path(__file__).parents[1]
def main():
 p=argparse.ArgumentParser(); p.add_argument('source',type=Path); p.add_argument('--output',type=Path,default=ROOT/'examples'/'GIP_Navigator_EXAMPLE.xlsx'); a=p.parse_args()
 d=preview(a.source.resolve()); a.output.parent.mkdir(parents=True,exist_ok=True)
 w=xlsxwriter.Workbook(a.output); w.set_calc_mode('auto')
 blue='#125AB6'; purple='#7B3F98'; orange='#F5A623'; cyan='#0C9BD6'; green='#00B050'; pale='#F4F7FB'; yellow='#FFF2CC'
 title=w.add_format({'bold':1,'font_color':'white','font_size':16,'bg_color':blue,'align':'center','valign':'vcenter','text_wrap':1})
 hdr=w.add_format({'bold':1,'font_color':'white','bg_color':purple,'align':'center','valign':'vcenter','text_wrap':1,'border':1})
 text=w.add_format({'font_name':'Aptos','font_size':10,'text_wrap':1,'valign':'top','border':1,'border_color':'#E5E7EB'})
 inp=w.add_format({'font_name':'Aptos','font_size':10,'text_wrap':1,'valign':'top','bg_color':yellow,'border':1,'border_color':'#E5E7EB'})
 dt=w.add_format({'num_format':'dd.mm.yyyy','font_name':'Aptos','font_size':10,'border':1,'border_color':'#E5E7EB'})
 dinp=w.add_format({'num_format':'dd.mm.yyyy','font_name':'Aptos','font_size':10,'bg_color':yellow,'border':1,'border_color':'#E5E7EB'})
 money=w.add_format({'num_format':'#,##0.00','font_name':'Aptos','font_size':10,'border':1,'border_color':'#E5E7EB'})
 note=w.add_format({'italic':1,'font_color':'#64748B','bg_color':pale,'text_wrap':1,'valign':'vcenter'})
 card=lambda c:w.add_format({'bold':1,'font_color':'white','bg_color':c,'align':'center','valign':'vcenter','text_wrap':1})
 home=w.add_worksheet('00_Главная'); kp=w.add_worksheet('01_КП'); lib=w.add_worksheet('02_Библиотека_СР'); sr=w.add_worksheet('03_СР_проекта'); agr=w.add_worksheet('04_Согласования'); g=w.add_worksheet('05_График'); ctl=w.add_worksheet('06_Контроль'); acts=w.add_worksheet('07_Акты')
 for s,c in [(home,blue),(kp,purple),(lib,cyan),(sr,orange),(agr,cyan),(g,purple),(ctl,orange),(acts,green)]: s.set_tab_color(c); s.hide_gridlines(2)
 # Dashboard
 home.set_column('A:P',13); home.merge_range('A1:P2','GIP Navigator',title); home.set_row(0,28);home.set_row(1,28)
 cards=[('A4:D5',f'ПРОЕКТ\n{d["project"]}',purple),('E4:H5',f'ЭТАПОВ КП\n{len(d["stages"])}',cyan),('I4:L5',f'СТОИМОСТЬ КП\n{float(d["total_rub_without_vat"]):,.2f} ₽',orange),('M4:P5','СТАТУС\nКП импортировано',green)]
 for rng,val,col in cards: home.merge_range(rng,val,card(col))
 home.merge_range('A8:P9','Дальнейшее заполнение: на листе «03_СР_проекта» выберите этап КП и внесите фактические даты. Жёлтые ячейки редактируются.',note)
 home.set_row(7,26); home.set_column('A:A',18)
 # КП
 kp.set_column('A:A',31);kp.set_column('B:B',42);kp.set_column('C:C',11);kp.set_column('D:D',56);kp.set_column('E:F',15);kp.set_column('G:G',23)
 kp.merge_range('A1:G1',f'Календарный план — {d["project"]}',title);kp.set_row(0,30);kp.merge_range('A2:G2',f'Источник: {a.source.name} · исходный файл не изменялся · импортировано этапов: {len(d["stages"])}',note)
 heads=['Номер договора / Номер Наряд-Заказа','Наименование проекта','№ Этапа','Наименование этапа','Начало работ','Окончание работ','Стоимость этапа, руб. без НДС'];kp.write_row('A3',heads,hdr);kp.set_row(2,34)
 for r,x in enumerate(d['stages'],3):
  kp.write(r,0,x['contract_or_order'],text);kp.write(r,1,x['project'],text);kp.write_string(r,2,x['stage'],text);kp.write(r,3,x['name'],text)
  (kp.write_datetime(r,4,x['start'],dt) if x['start'] else kp.write_blank(r,4,None,dt));(kp.write_datetime(r,5,x['end'],dt) if x['end'] else kp.write_blank(r,5,None,dt));(kp.write_number(r,6,float(x['amount_rub_without_vat']),money) if x['amount_rub_without_vat'] is not None else kp.write_blank(r,6,None,money)); kp.set_row(r,36)
 kp.autofilter(2,0,2+len(d['stages']),6);kp.freeze_panes(3,0)
 # library
 lib.set_column('A:A',12);lib.set_column('B:B',42);lib.set_column('C:C',12);lib.set_column('D:D',14);lib.set_column('E:E',32);lib.merge_range('A1:E1','Библиотека сопутствующих работ (СР)',title);lib.merge_range('A2:E2','Базовые позиции можно редактировать. При необходимости добавьте строки и перенесите их в план СР проекта.',note);lib.write_row('A3',['Код СР','Наименование работы','Длительность, дн.','Связь','Комментарий'],hdr)
 base=[('SR-01','Получение исходных данных',3,'FS','Стартовая работа'),('SR-02','Инженерные изыскания',20,'FS','По графику изысканий'),('SR-03','Разработка проектной документации',30,'FS','Основной комплект'),('SR-04','Внутренняя проверка',5,'FS','Контроль качества'),('SR-05','Согласование с заказчиком',10,'FS','Срок зависит от замечаний'),('SR-06','Экспертиза',30,'FS','При необходимости'),('SR-07','Рабочая документация',25,'FS','После утверждения'),('SR-08','Передача комплекта',3,'FS','Финальная выдача')]
 for r,row in enumerate(base,3):
  for c,val in enumerate(row): lib.write(r,c,val,inp if c in (2,3) else text)
  lib.set_row(r,26)
 lib.freeze_panes(3,0)
 # SR project
 sr.set_column('A:A',12);sr.set_column('B:B',38);sr.set_column('C:C',12);sr.set_column('D:E',15);sr.set_column('F:F',11);sr.set_column('G:G',15);sr.set_column('H:I',16);sr.set_column('J:J',22)
 sr.merge_range('A1:J1','План сопутствующих работ (СР)',title);sr.merge_range('A2:J2','Жёлтые ячейки — ввод. Выберите этап КП, затем задайте плановую дату и фактические/ожидаемые даты.',note)
 srh=['Код СР','Работа','Этап КП','План начала','Длит., дн.','План окончания','Факт. окончание','Ожид. окончание','Прогноз окончания','Статус'];sr.write_row('A3',srh,hdr)
 for r,row in enumerate(base,3):
  excel=r+1;sr.write(r,0,row[0],text);sr.write(r,1,row[1],text);sr.write_blank(r,2,None,inp);sr.write_blank(r,3,None,dinp);sr.write_number(r,4,row[2],inp);sr.write_formula(r,5,f'=IF(OR(D{excel}="",E{excel}=""),"",D{excel}+E{excel}-1)',dt);sr.write_blank(r,6,None,dinp);sr.write_blank(r,7,None,dinp);sr.write_formula(r,8,f'=IF(G{excel}<>"",G{excel},IF(H{excel}<>"",H{excel},F{excel}))',dt);sr.write_formula(r,9,f'=IF(C{excel}="","Не задан этап КП",IF(I{excel}="","Нет срока СР",IFERROR(IF(I{excel}>VLOOKUP(C{excel},\'01_КП\'!$C$4:$F${len(d["stages"])+3},4,FALSE),"Критично: позднее КП","Норма"),"Нет срока КП")))',text);sr.set_row(r,27)
 sr.data_validation(3,2,10,2,{'validate':'list','source':f"='01_КП'!$C$4:$C${len(d['stages'])+3}"});sr.conditional_format(3,9,10,9,{'type':'text','criteria':'containing','value':'Критично','format':w.add_format({'bg_color':'#FFC7CE','font_color':'#9C0006'})});sr.freeze_panes(3,0)
 # agreements and acts
 for ws,titletext,heads,widths in [(agr,'Согласования',['Документ / вопрос','Ответственный','Дата направления','Срок ответа','Статус','Комментарий'],[34,24,15,15,18,42]),(acts,'Акты и выдача',['Документ','Этап КП','Дата выдачи','Дата подписания','Статус','Комментарий'],[34,14,15,17,18,42])]:
  for col,width in enumerate(widths):ws.set_column(col,col,width)
  ws.merge_range(0,0,0,len(heads)-1,titletext,title);ws.merge_range(1,0,1,len(heads)-1,'Заполняйте фактические сведения. Пустые строки предназначены для новых записей.',note);ws.write_row(2,0,heads,hdr)
  for r in range(3,23):ws.write_blank(r,0,None,inp);[ws.write_blank(r,c,None,dinp if c in (2,3) else inp) for c in range(1,len(heads))];ws.set_row(r,24)
  ws.freeze_panes(3,0)
 # Gantt source and chart
 g.set_column('A:A',9);g.set_column('B:B',52);g.set_column('C:C',14);g.set_column('D:F',14);g.write_row('A1',['Тип','Наименование','Начало','Длительность','КП — срок','СР — срок'],hdr)
 gr=1
 for x in d['stages']:
  if x['start'] and x['end']:
   g.write(gr,0,'КП',text);g.write(gr,1,x['name'],text);g.write_datetime(gr,2,x['start'],dt);g.write_number(gr,3,(x['end']-x['start']).days+1,text);g.write_number(gr,4,(x['end']-x['start']).days+1,text);g.write_formula(gr,5,'=NA()',text);gr+=1
 for i,row in enumerate(base,4):
  g.write(gr,0,'СР',text);g.write_formula(gr,1,f"='03_СР_проекта'!B{i}",text);g.write_formula(gr,2,f"='03_СР_проекта'!D{i}",dt);g.write_formula(gr,3,f"=IF(I{i}=\"\",NA(),I{i}-C{gr+1}+1)",text);g.write_formula(gr,4,'=NA()',text);g.write_formula(gr,5,f'=IF(D{gr+1}=NA(),NA(),D{gr+1})',text);gr+=1
 chart=w.add_chart({'type':'bar','subtype':'stacked'});end=gr
 chart.add_series({'categories':f"='05_График'!$B$2:$B${end}",'values':f"='05_График'!$C$2:$C${end}",'fill':{'none':True},'border':{'none':True}});chart.add_series({'name':'КП — договорные сроки','categories':f"='05_График'!$B$2:$B${end}",'values':f"='05_График'!$E$2:$E${end}",'fill':{'color':purple}});chart.add_series({'name':'СР — рабочие сроки','categories':f"='05_График'!$B$2:$B${end}",'values':f"='05_График'!$F$2:$F${end}",'fill':{'color':orange}});chart.set_title({'name':'Диаграмма Ганта: КП и СР'});chart.set_legend({'position':'bottom'});chart.set_y_axis({'reverse':True});chart.set_size({'width':1280,'height':max(600,120+22*(end-1))});g.insert_chart('A31',chart);g.freeze_panes(1,0)
 # control
 ctl.set_column('A:A',12);ctl.set_column('B:B',38);ctl.set_column('C:C',16);ctl.set_column('D:D',28);ctl.merge_range('A1:D1','Контроль сроков СР',title);ctl.write_row('A3',['Код СР','Работа','Прогноз','Сигнал'],hdr)
 for r in range(3,11):
  er=r+1;ctl.write_formula(r,0,f"='03_СР_проекта'!A{er}",text);ctl.write_formula(r,1,f"='03_СР_проекта'!B{er}",text);ctl.write_formula(r,2,f"='03_СР_проекта'!I{er}",dt);ctl.write_formula(r,3,f"='03_СР_проекта'!J{er}",text)
 ctl.conditional_format(3,3,10,3,{'type':'text','criteria':'containing','value':'Критично','format':w.add_format({'bg_color':'#FFC7CE','font_color':'#9C0006'})})
 w.close();print(a.output)
if __name__=='__main__':main()
