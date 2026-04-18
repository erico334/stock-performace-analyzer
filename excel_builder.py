"""
excel_builder.py - Uses xlsxwriter 
"""
import pandas as pd
import numpy as np
from io import BytesIO
import xlsxwriter
import analyzer as az

DARK    = '#1F3864'
MID_BLU = '#2E75B6'
TOT_BG  = '#D6DCE4'
WHITE   = '#FFFFFF'
FONT    = 'Arial'

ROW_COLORS = {
    'fast':  ('#E2EFDA', '#EFFAF0'),
    'active':('#EBF5E0', '#F5FCF0'),
    'recent':('#F2F8EC', '#FAFFF5'),
    'slow1': ('#FFF2CC', '#FFFDF0'),
    'slow2': ('#FFEB9C', '#FFFDE8'),
    'dorm1': ('#FCE4D6', '#FDF0E8'),
    'dorm2': ('#F8CBAD', '#FDDCCA'),
    'dead1': ('#F4CCCC', '#FAE0E0'),
    'dead2': ('#EAAAA0', '#F5C6C0'),
    'dead3': ('#D9534F', '#D9534F'),
    'never': ('#F2F2F2', '#FFFFFF'),
}

def _row_color(days, idx):
    alt = idx % 2
    if pd.isna(days): return ROW_COLORS['never'][alt]
    d = int(days)
    if d >= 366: return ROW_COLORS['dead3'][0]
    if d >= 271: return ROW_COLORS['dead2'][alt]
    if d >= 181: return ROW_COLORS['dead1'][alt]
    if d >= 121: return ROW_COLORS['dorm2'][alt]
    if d >= 91:  return ROW_COLORS['dorm1'][alt]
    if d >= 61:  return ROW_COLORS['slow2'][alt]
    if d >= 31:  return ROW_COLORS['slow1'][alt]
    if d >= 15:  return ROW_COLORS['recent'][alt]
    if d >= 8:   return ROW_COLORS['active'][alt]
    return ROW_COLORS['fast'][alt]

def _flag(row):
    days  = row.get('DAYS_SINCE_SALE')
    qty_h = row.get('QTY AT HAND', 0)
    qty_s = row.get('QTY SOLD', 0)
    if qty_h > 0 and qty_h <= 5 and qty_s > 100:
        return 'LOW STOCK - reorder!'
    if pd.isna(days): return 'Never sold'
    d = int(days)
    if d >= 366: return 'Write-off risk'
    if d >= 271: return 'Critical - liquidate'
    if d >= 181: return 'Near dead'
    if d >= 121: return 'Deeply dormant'
    if d >= 91:  return 'Dormant'
    if d >= 61:  return 'At risk - act now'
    if d >= 31:  return 'Slowing - monitor'
    if d >= 15:  return 'Recent sale'
    if d >= 8:   return 'Active'
    return 'Fast mover'


class ReportBuilder:
    def __init__(self, df, metrics):
        self.df      = df
        self.metrics = metrics
        self.buf     = BytesIO()
        self.wb      = xlsxwriter.Workbook(self.buf, {'in_memory': True, 'nan_inf_to_errors': True})
        self._init_formats()

    def _init_formats(self):
        wb = self.wb
        def fmt(opts):
            base = {'font_name': FONT, 'font_size': 10, 'border': 1, 'border_color': '#CCCCCC'}
            base.update(opts)
            return wb.add_format(base)
        self.f = {
            'title':    fmt({'bold':True,'font_size':13,'font_color':WHITE,'bg_color':DARK,'align':'left','valign':'vcenter'}),
            'subtitle': fmt({'italic':True,'font_size':10,'font_color':WHITE,'bg_color':'#2E4057','align':'left','valign':'vcenter'}),
            'hdr':      fmt({'bold':True,'font_color':WHITE,'bg_color':DARK,'align':'center','valign':'vcenter'}),
            'hdr_red':  fmt({'bold':True,'font_color':WHITE,'bg_color':'#C00000','align':'center','valign':'vcenter'}),
            'hdr_gray': fmt({'bold':True,'font_color':WHITE,'bg_color':'#595959','align':'center','valign':'vcenter'}),
            'hdr_blue': fmt({'bold':True,'font_color':WHITE,'bg_color':MID_BLU,'align':'left','valign':'vcenter'}),
            'card_lbl': fmt({'bold':True,'font_size':9,'font_color':'#595959','bg_color':'#F5F5F5','align':'center','valign':'vcenter'}),
            'card_val': fmt({'bold':True,'font_size':13,'font_color':DARK,'bg_color':WHITE,'align':'center','valign':'vcenter','num_format':'#,##0'}),
            'total':    fmt({'bold':True,'font_color':'#1F1F1F','bg_color':TOT_BG,'align':'center','valign':'vcenter','num_format':'#,##0'}),
            'total_lbl':fmt({'bold':True,'font_color':'#1F1F1F','bg_color':TOT_BG,'align':'left','valign':'vcenter'}),
        }

    def _cf(self, bg, extra=None):
        opts = {'font_name':FONT,'font_size':9,'border':1,'border_color':'#CCCCCC',
                'bg_color':bg,'valign':'vcenter','align':'center'}
        if extra: opts.update(extra)
        return self.wb.add_format(opts)

    def _write_title(self, ws, text, subtitle, col_end=8):
        ws.merge_range(0,1,0,col_end, text,    self.f['title'])
        ws.merge_range(1,1,1,col_end, subtitle, self.f['subtitle'])
        ws.set_row(0,28); ws.set_row(1,18); ws.set_row(2,8)

    def _write_metric_cards(self, ws, cards, start_row=3):
        for i,(lbl,val) in enumerate(cards):
            sc=i*2+1; ec=sc+1
            ws.merge_range(start_row,  sc,start_row,  ec, lbl, self.f['card_lbl'])
            ws.merge_range(start_row+1,sc,start_row+1,ec, val, self.f['card_val'])
        ws.set_row(start_row,18); ws.set_row(start_row+1,26); ws.set_row(start_row+2,8)

    def _write_headers(self, ws, headers, row, fmt_key='hdr'):
        fmt = self.f[fmt_key]
        for i,h in enumerate(headers,1):
            ws.write(row,i,h,fmt)
        ws.set_row(row,20)

    def build_summary(self):
        m  = self.metrics
        ws = self.wb.add_worksheet('Summary Dashboard')
        ws.hide_gridlines(2)
        ws.set_column(0,0,3); ws.set_column(1,1,28); ws.set_column(2,2,12)
        ws.set_column(3,3,12); ws.set_column(4,4,14); ws.set_column(5,5,16)
        ws.set_column(6,6,18); ws.set_column(7,7,18); ws.set_column(8,8,26)
        snap = m['snapshot_date'].strftime('%d/%m/%Y')
        self._write_title(ws,
            'STOCK PERFORMANCE - COMPREHENSIVE ANALYSIS REPORT',
            f"Snapshot: {snap}  |  Total SKUs: {m['total_skus']:,}  |  Stock Performance Analyzer")
        self._write_metric_cards(ws,[
            ('Total SKUs',       m['total_skus']),
            ('Active (<=30d)',   m['active']),
            ('Never Sold',       m['never_sold']),
            ('Total Revenue (N)',int(m['total_revenue'])),
            ('Capital at Risk',  int(m['idle_capital'])),
        ],start_row=3)
        ws.merge_range(6,1,6,8,'STOCK HEALTH BY AGE BUCKET',self.f['hdr_blue'])
        ws.set_row(6,22)
        self._write_headers(ws,['Age Bucket','Category','Total SKUs','With Stock',
            'Units in Stock','Capital Tied (N)','Avg Days Idle','Risk Level','Recommended Action'],7)
        bucket_df = az.get_bucket_summary(self.df)
        risk_bgs = {'Low':'#E2EFDA','Medium':'#FFF2CC','High':'#FCE4D6','Critical':'#F4CCCC'}
        risk_tcs = {'Low':'#375623','Medium':'#7D5A00','High':'#843C0C','Critical':'#7F0000'}
        for ri,row in bucket_df.iterrows():
            r=ri+8; bg=row['_bg']; rl=row['Risk Level']
            ws.write(r,1,row['Age Bucket'],      self._cf(bg,{'align':'left','bold':True,'font_color':row['_tc'].lstrip('#')}))
            ws.write(r,2,row['Category'],        self._cf(bg))
            ws.write(r,3,row['Total SKUs'],      self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,4,row['With Stock'],      self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,5,row['Units in Stock'],  self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,6,row['Capital Tied (NGN)'],self._cf(bg,{'num_format':'#,##0'}))
            avg=row['Avg Days Idle']
            ws.write(r,7,round(float(avg),1) if (avg and not pd.isna(avg)) else 0,self._cf(bg,{'num_format':'0.0'}))
            ws.write(r,8,rl,self._cf(risk_bgs.get(rl,'#F2F2F2'),{'bold':True,'font_color':risk_tcs.get(rl,'#1F1F1F')}))
            ws.write(r,9,row['Action'],          self._cf(bg,{'align':'left'}))
            ws.set_row(r,20)
        tr=len(bucket_df)+8
        ws.write(tr,1,'TOTAL',self.f['total_lbl'])
        for ci in range(2,9): ws.write(tr,ci,'',self.f['total'])
        ws.set_row(tr,22)

    def build_detail(self, sheet_name, title, subtitle, src_df):
        ws = self.wb.add_worksheet(sheet_name[:31])
        ws.hide_gridlines(2)
        ws.set_column(0,0,3); ws.set_column(1,1,6); ws.set_column(2,2,44)
        ws.set_column(3,3,18); ws.set_column(4,4,13); ws.set_column(5,5,13)
        ws.set_column(6,6,14); ws.set_column(7,7,16); ws.set_column(8,8,18)
        ws.set_column(9,9,14); ws.set_column(10,10,22)
        self._write_title(ws,title,subtitle,col_end=10)
        ws_df=src_df[src_df['QTY AT HAND']>0]
        avg_d=round(src_df['DAYS_SINCE_SALE'].mean(),1) if src_df['DAYS_SINCE_SALE'].notna().any() else 0
        self._write_metric_cards(ws,[
            ('Total SKUs',        len(src_df)),
            ('With Stock',        len(ws_df)),
            ('Units in Stock',    int(ws_df['QTY AT HAND'].sum())),
            ('Capital Tied (N)',  int(ws_df['CAPITAL_TIED'].sum())),
            ('Avg Days Since Sale',round(avg_d,1)),
        ],start_row=3)
        self._write_headers(ws,['#','Item Name','Last Sale Date','Days Idle','Qty Sold',
            'Qty At Hand','Unit Cost (N)','Capital Tied (N)','Barcode','Status / Flag'],6)
        sdf=src_df.copy()
        sdf['_s']=sdf['DAYS_SINCE_SALE'].fillna(99999)
        sdf=sdf.sort_values(['_s','CAPITAL_TIED'],ascending=[False,False]).reset_index(drop=True)
        for idx,row in sdf.iterrows():
            r=idx+7; days=row['DAYS_SINCE_SALE']; qty_h=int(row['QTY AT HAND'])
            bg=_row_color(days,idx); flag=_flag(row)
            ws.write(r,1, idx+1,                                         self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,2, row['ITEM NAME'],                              self._cf(bg,{'align':'left'}))
            ws.write(r,3, row['LAST_SALE_STR'],                          self._cf(bg))
            ws.write(r,4, int(days) if pd.notna(days) else '---',       self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,5, int(row['QTY SOLD']),                         self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,6, qty_h,                                         self._cf(bg,{'num_format':'#,##0','font_color':'#C00000' if qty_h<0 else '#1F1F1F','bold':qty_h<0}))
            ws.write(r,7, row['UNIT_COST'] if row['UNIT_COST']>0 else 0,self._cf(bg,{'num_format':'#,##0.00'}))
            cap=row['CAPITAL_TIED'] if row['CAPITAL_TIED']>0 else 0
            ws.write(r,8, cap,                                           self._cf(bg,{'num_format':'#,##0','font_color':'#843C0C' if cap>100000 else '#1F1F1F'}))
            ws.write(r,9, row.get('BARCODE_STR',''),                     self._cf(bg))
            ws.write(r,10,flag,                                          self._cf(bg,{'align':'left','font_color':'#C00000' if 'LOW STOCK' in flag else '#1F1F1F'}))
            ws.set_row(r,15)
        tr=len(sdf)+7
        ws.write(tr,1,'TOTAL',self.f['total_lbl'])
        for ci in range(2,11): ws.write(tr,ci,'',self.f['total'])
        ws.set_row(tr,20); ws.freeze_panes(7,2)

    def build_monthly(self):
        ws=self.wb.add_worksheet('Monthly Trend')
        ws.hide_gridlines(2)
        ws.set_column(0,0,3); ws.set_column(1,1,6); ws.set_column(2,2,20)
        ws.set_column(3,3,16); ws.set_column(4,4,16); ws.set_column(5,5,18)
        self._write_title(ws,'MONTHLY SALES TREND','Products by last-sale month',col_end=5)
        self._write_headers(ws,['Month','Active SKUs','Units Sold','Revenue (N)','Avg Rev/SKU'],3)
        trend=az.get_monthly_trend(self.df)
        bgs=['#EBF5E0','#FFF2CC','#FCE4D6','#D9E8F5','#E8D5F5','#F2F8EC','#FFEB9C','#F8CBAD']
        for ri,row in trend.iterrows():
            r=ri+4; bg=bgs[ri%len(bgs)]; rps=row['revenue']/row['skus'] if row['skus']>0 else 0
            ws.write(r,1,row['MONTH_STR'],    self._cf(bg,{'bold':True}))
            ws.write(r,2,int(row['skus']),    self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,3,int(row['qty_sold']),self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,4,row['revenue'],      self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,5,rps,                 self._cf(bg,{'num_format':'#,##0'}))
            ws.set_row(r,20)
        tr=len(trend)+4
        ws.write(tr,1,'TOTAL',self.f['total_lbl'])
        for ci in range(2,6): ws.write(tr,ci,'',self.f['total'])
        ws.set_row(tr,20)

    def build_top_products(self):
        ws=self.wb.add_worksheet('Top Products')
        ws.hide_gridlines(2)
        ws.set_column(0,0,3); ws.set_column(1,1,6); ws.set_column(2,2,44)
        ws.set_column(3,3,14); ws.set_column(4,4,13); ws.set_column(5,5,14)
        ws.set_column(6,6,16); ws.set_column(7,7,18); ws.set_column(8,8,20)
        self._write_title(ws,'TOP 50 PRODUCTS - BY REVENUE','Ranked by cost-basis revenue. Red = low stock.',col_end=8)
        self._write_headers(ws,['Rank','Item Name','Qty Sold','Qty At Hand','Unit Cost (N)','Revenue (N)','Last Sale Date','Status'],2)
        top50=az.get_top_products(self.df,n=50,by='REVENUE')
        for ri,row in top50.iterrows():
            r=ri+3; qty_h=int(row.get('QTY AT HAND',0)); low=qty_h<=5 and row['QTY SOLD']>100
            bg='#FFE0E0' if low else ('#E2EFDA' if ri%2==0 else WHITE)
            ws.write(r,1,ri+1,                        self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,2,row['ITEM NAME'],             self._cf(bg,{'align':'left'}))
            ws.write(r,3,int(row['QTY SOLD']),         self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,4,qty_h,                        self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,5,row.get('UNIT_COST',0),       self._cf(bg,{'num_format':'#,##0.00'}))
            ws.write(r,6,row['REVENUE'],               self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,7,row.get('LAST_SALE_STR',''),  self._cf(bg))
            ws.write(r,8,'LOW STOCK' if low else row.get('STATUS',''),self._cf(bg,{'align':'left'}))
            ws.set_row(r,15)
        ws.freeze_panes(3,2)

    def build_dead_stock(self):
        dead=az.get_dead_stock(self.df)
        ws=self.wb.add_worksheet('Dead Stock (Never Sold)')
        ws.hide_gridlines(2)
        ws.set_column(0,0,3); ws.set_column(1,1,6); ws.set_column(2,2,44)
        ws.set_column(3,3,13); ws.set_column(4,4,15); ws.set_column(5,5,17)
        ws.set_column(6,6,15); ws.set_column(7,7,24)
        self._write_title(ws,'DEAD STOCK - ZERO SALES, HAS PHYSICAL INVENTORY',
            'Products with QTY AT HAND > 0 that have never recorded a sale.',col_end=7)
        ws_df=dead[dead['QTY AT HAND']>0]
        self._write_metric_cards(ws,[
            ('Dead SKUs',       len(ws_df)),
            ('Total Units',     int(ws_df['QTY AT HAND'].sum())),
            ('Capital Tied (N)',int(ws_df['CAPITAL_TIED'].sum())),
            ('Avg Unit Cost',   int(ws_df['UNIT_COST'].mean()) if len(ws_df)>0 else 0),
        ],start_row=3)
        self._write_headers(ws,['#','Item Name','Qty At Hand','Unit Cost (N)',
            'Capital Tied (N)','Barcode','Recommendation'],6,fmt_key='hdr_gray')
        for ri,row in dead.iterrows():
            r=ri+7; cap=row['CAPITAL_TIED']
            bg='#FFE0B2' if cap>50000 else ('#FFF9E6' if cap>10000 else ('#F2F2F2' if ri%2==0 else WHITE))
            ws.write(r,1,ri+1,                                         self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,2,row['ITEM NAME'],                             self._cf(bg,{'align':'left'}))
            ws.write(r,3,int(row['QTY AT HAND']),                      self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,4,row['UNIT_COST'] if row['UNIT_COST']>0 else 0,self._cf(bg,{'num_format':'#,##0.00'}))
            ws.write(r,5,cap,                                           self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,6,row.get('BARCODE_STR',''),                     self._cf(bg))
            ws.write(r,7,row.get('RECOMMENDATION','Clear or de-list'),  self._cf(bg,{'align':'left'}))
            ws.set_row(r,15)
        tr=len(dead)+7
        ws.write(tr,1,'TOTAL',self.f['total_lbl'])
        for ci in range(2,8): ws.write(tr,ci,'',self.f['total'])
        ws.set_row(tr,20); ws.freeze_panes(7,2)

    def build_negative_stock(self):
        neg=az.get_negative_stock(self.df)
        if len(neg)==0: return
        ws=self.wb.add_worksheet('Negative Stock (Data Issue)')
        ws.hide_gridlines(2)
        ws.set_column(0,0,3); ws.set_column(1,1,6); ws.set_column(2,2,44)
        ws.set_column(3,3,14); ws.set_column(4,4,13); ws.set_column(5,5,15); ws.set_column(6,6,28)
        self._write_title(ws,'NEGATIVE INVENTORY - DATA INTEGRITY ALERT',
            f'{len(neg)} products show negative stock. Likely bulk-break/sachet tracking errors.',col_end=6)
        self._write_headers(ws,['#','Item Name','Qty At Hand','Qty Sold','Unit Cost (N)','Likely Cause'],3,fmt_key='hdr_red')
        for ri,row in neg.iterrows():
            r=ri+4; bg='#FFE0E0' if ri%2==0 else '#FFF0F0'
            ws.write(r,1,ri+1,                                          self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,2,row['ITEM NAME'],                              self._cf(bg,{'align':'left'}))
            ws.write(r,3,int(row['QTY AT HAND']),                       self._cf(bg,{'num_format':'#,##0','font_color':'#C00000','bold':True}))
            ws.write(r,4,int(row['QTY SOLD']),                          self._cf(bg,{'num_format':'#,##0'}))
            ws.write(r,5,row['UNIT_COST'] if row['UNIT_COST']>0 else 0, self._cf(bg,{'num_format':'#,##0.00'}))
            ws.write(r,6,'Bulk-break / sachet tracking mismatch',        self._cf(bg,{'align':'left'}))
            ws.set_row(r,15)
        ws.freeze_panes(4,2)

    def build(self):
        df=self.df
        self.build_summary()
        slow=df[df['DAYS_SINCE_SALE'].between(31,90,inclusive='both')].copy()
        self.build_detail('Slow Moving (31-90 Days)','SLOW MOVING - Last sale 31 to 90 days ago',
            f"{len(slow)} products | {(slow['QTY AT HAND']>0).sum()} with stock",slow)
        dorm=df[df['DAYS_SINCE_SALE'].between(91,180,inclusive='both')].copy()
        self.build_detail('Dormant (91-180 Days)','DORMANT - Last sale 91 to 180 days ago',
            f"{len(dorm)} products | {(dorm['QTY AT HAND']>0).sum()} with stock",dorm)
        nd=df[df['DAYS_SINCE_SALE']>=181].copy()
        self.build_detail('Near Dead (181+ Days)','NEAR DEAD - Last sale 181+ days ago',
            f"{len(nd)} products | {(nd['QTY AT HAND']>0).sum()} still have stock",nd)
        self.build_dead_stock()
        all_stock=az.get_all_stock_by_idle(df)
        self.build_detail('All Stock - By Idle Days','ALL IN-STOCK PRODUCTS - Ranked oldest last-sale first',
            f"{len(all_stock)} products with QTY AT HAND > 0",all_stock)
        self.build_top_products()
        self.build_monthly()
        self.build_negative_stock()
        sold=df[df['LAST SALES DATE'].notna()].copy()
        self.build_detail('All Sold - Chronological','ALL SOLD PRODUCTS - Full historical register oldest first',
            f"{len(sold)} products that have ever recorded a sale",sold)
        self.wb.close()
        self.buf.seek(0)
        return self.buf.getvalue()


def build_report(df, metrics):
    return ReportBuilder(df, metrics).build()
