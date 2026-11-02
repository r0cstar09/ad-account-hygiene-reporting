#!/usr/bin/env python3
"""Generate a deterministic, entirely fictional AD export. Never reads real data."""
from __future__ import annotations
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

HEADERS = ['Disabled', 'Display Name', 'SAM Account Name', 'Distinguished Name', 'Canonical Name', 'Account Expires On', 'Password Last Set', 'When Created', 'Last Logon Timestamp', 'Days Since Last Logon', 'Logon Count']
AS_OF = datetime(2026, 11, 1, 12, 0, 0)
GROUPS = [
    ('Workforce', 'Users/Region-A/Accounts/Site-A', ['Never Logged On', 810, 400, 230, 190, 105, 95, 83, '', 'unavailable']),
    ('Privileged', 'Privileged/Administrators/Operations', ['Never Logged On', 455, 190, 120, 85]),
    ('Service', 'Applications/Service Accounts/Batch', ['Never Logged On', 380, 210, 100, 'Never Logged On', 500]),
    ('Other', 'Unclassified/Contractors/Review', ['Never Logged On', 205, 115]),
]

def make_report(output: Path) -> int:
    wb=Workbook(); ws=wb.active; assert ws is not None; ws.title='Report'
    ws['A1']='example.test | SYNTHETIC DEMO';ws['K1']=AS_OF;ws['K1'].number_format='dd/mm/yyyy hh:mm:ss'
    ws.merge_cells('A2:K2');ws['A2']='Enabled Users who Never Logged on or have Not Logged on for at least 90 days - SYNTHETIC'
    ws['A2'].font=Font(name='Calibri',size=14,bold=True,color='191970');ws['A2'].fill=PatternFill('solid',fgColor='D3D3D3');ws['A2'].alignment=Alignment(horizontal='center')
    for col,header in enumerate(HEADERS,1):
        c=ws.cell(3,col,header);c.font=Font(name='Calibri',size=12,bold=True,color='FFFFFF');c.fill=PatternFill('solid',fgColor='4682B4')
    serial=0
    for group,path,ages in GROUPS:
        for i,days in enumerate(ages,1):
            serial+=1
            name=f'Demo {group} {i:02d}'
            sam=f'demo.{group.lower()}.{i:02d}' if group!='Service' else f'svc_demo_{i:02d}'
            dn='CN='+name+','+','.join('OU='+p for p in reversed(path.split('/')))+',DC=example,DC=test'
            if isinstance(days,int):
                stamp=(AS_OF-timedelta(days=max(days+7,95))).strftime('%Y-%m-%d %H:%M:%S');count=serial*17
            elif days=='Never Logged On':stamp='Never';count=0
            else:stamp=(AS_OF-timedelta(days=120)).strftime('%Y-%m-%d %H:%M:%S');count=3
            created=(AS_OF-timedelta(days=14 if days=='Never Logged On' else 1000)).strftime('%Y-%m-%d %H:%M:%S')
            values=['',name,sam,dn,'example.test/'+path+'/'+name,'Never','' if days=='Never Logged On' else '2024-01-15 09:00:00',created,stamp,str(days),str(count)]
            for col,value in enumerate(values,1):
                c=ws.cell(serial+3,col,value);c.data_type='s';c.font=Font(name='Calibri',size=11);c.fill=PatternFill('solid',fgColor='F0F8FF');c.alignment=Alignment(vertical='center')
    widths={'A':12,'B':27,'C':24,'D':100,'E':90,'F':20,'G':23,'H':23,'I':25,'J':24,'K':15}
    for col,width in widths.items():ws.column_dimensions[col].width=width
    ws.freeze_panes='A4';ws.row_dimensions[2].height=20;ws.row_dimensions[3].height=18
    wb.properties.creator='Synthetic demonstration';wb.properties.title='Fictional AD data for testing and screenshots';wb.properties.description='No real organization, person, account or operational counts.'
    output.parent.mkdir(parents=True,exist_ok=True);wb.save(output)
    return serial

if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--output',type=Path,default=Path(__file__).resolve().parents[1]/'output'/'synthetic-input.xlsx');args=parser.parse_args()
    count=make_report(args.output);print(f'synthetic_rows={count}');print(f'output={args.output}')
