function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('AutoFill Docs');
  menu.addItem('Create New Docs', 'createNewGoogleDocs')
  menu.addToUi();
}

function checkSymbol (symbolValues) {
  const str = symbolValues.toString()
  return str.includes('<') || str.includes('≤')
}


function createNewGoogleDocs() {
    const googleDocTemplate = DriveApp.getFileById('1mQloNFvuenNh8-cF5m6zuG8zNa2DGdbuW4-EAJxkr3E');
    const destinationFolder = DriveApp.getFolderById('1jcs_cgbTO48-LY4gbsH5EnvLqoRpZSuo');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Supplier');
    const rows = sheet.getDataRange().getValues();
    // @ts-ignore
    let columnNames = [];

    rows.forEach(function(row, index){
      if (index === 0) return;
      if (index === 1) return;
      if (index === 2) {
        columnNames = row // ['no', 'nomor kontrak']
      };
      if (!row[1]) return;
      if (row[columnNames.indexOf('Document Link')]) return;
      
      const copy = googleDocTemplate.makeCopy(`${row[1]} Appendix` , destinationFolder)
      const doc = DocumentApp.openById(copy.getId())
      const body = doc.getBody();

      columnNames.forEach((column, colIdx) => {

        if (column === 'Toleransi') {
          Logger.log('Toleransi ' + typeof row[colIdx])
          if(row[colIdx].toString().includes('%')) {
          body.replaceText('{{Toleransi}}', 'dengan selisih'+row[colIdx]+' dari total kuantitas dapat diterima');
          }
          else {body.replaceText('{{Toleransi}}', row[colIdx]);}
        }

        if (column === 'Email Supplier') {
          if(!row[colIdx]) {
          body.replaceText('Email: {{Email Supplier}}', '');
          }
          else {body.replaceText('{{Email Supplier}}', row[colIdx]);}
        }

        if (column === 'Harga') {
          if(!row[columnNames.indexOf('Nominal')]) {
            body.replaceText('{{Harga}}','HPM Bijih Nikel dengan Basis FOB.');
          }
            if(row[colIdx].toString().includes('HPM')) {
              body.replaceText('{{Harga}}','HPM Bijih Nikel dengan Basis FOB + '+row[columnNames.indexOf('Nominal')]+' USD/WMT yang akan dihitung bersamaan dengan pembayaran akhir Bijih Nikel dan merupakan satu kesatuan yang tidak dapat dipisahkan.');
            }
            else { body.replaceText('{{Harga}}',row[colIdx]+' '+row[columnNames.indexOf('Nominal')]+' USD/WMT.');
            } 
        }

        if (column === 'Incoterms') {
          if(row[colIdx]=='CIF') {
          body.replaceText('{{LoadDisch}}','Pembongkaran')
          }
          else {body.replaceText('{{LoadDisch}}','Pemuatan')}
        }

        if (column === 'Demurrage'){
          if(!row[colIdx]) {
          body.replaceText('Demurrage	: {{Demurrage}}','');
          body.replaceText('& Demurrage','');
          }
          else {body.replaceText('{{Demurrage}}', row[colIdx]);}
        }

        //SPECS DAN PENYESUAIAN DI BAWAH
        
        //Ni START
        //BONUS Ni
        if (column === 'Bonus Ni 1') {
          if (!row[colIdx]) {
          body.replaceText ('Ni {{Symbol Bonus Ni}}{{Bonus Ni}}','');
          body.replaceText ('per kenaikan {{Calc Bonus Ni}}','');
          body.replaceText ('\\(\\+\\) {{Harga Bonus Ni}} USD/WMT','');
          }
          else {
            body.replaceText('{{Symbol Bonus Ni}}', row[columnNames.indexOf('Symbol Bonus Ni 1')]);
            body.replaceText('{{Bonus Ni}}', row[columnNames.indexOf('Bonus Ni 1')]);
            body.replaceText('{{Calc Bonus Ni}}', row[columnNames.indexOf('Calc Bonus Ni 1')]);
            body.replaceText('{{Harga Bonus Ni}}', row[columnNames.indexOf('Harga Bonus Ni 1')]);
          };
        }
        //PENALTY Ni 1 
        if (column === 'Penalty Ni 1') {
          if (!row[colIdx]){
          body.replaceText ('Ni {{Symbol Ni 1}}{{Penalty Ni 1}}','');
          body.replaceText ('per penurunan {{Calc Ni 1}}','');
          body.replaceText ('\\(-\\) {{Harga Ni 1}} USD/WMT','');
          }
          else {body.replaceText('{{Symbol Ni 1}}', row[columnNames.indexOf('Penalty Ni 1')]);
                body.replaceText('{{Penalty Ni 1}}', row[columnNames.indexOf('Penalty Ni 1')]);
                if(checkSymbol(row[columnNames.indexOf('Symbol Ni 1')])) 
                  {body.replaceText('{{Pembagi Ni 1}}', 'per penurunan');}
                  else {body.replaceText('{{Pembagi Ni 1}}', 'per kenaikan');}
                body.replaceText('{{Calc Ni 1}}', row[columnNames.indexOf('Calc Ni 1')]);
                body.replaceText('{{Harga Ni 1}}', row[columnNames.indexOf('Harga Ni 1')]);
          };
        }
        //PENALTY Ni 2
        if (column === 'Penalty Ni 2') {
          if (!row[colIdx]) {
          body.replaceText ('Ni {{Symbol Ni 2}}{{Penalty Ni 2}} \\(Penalti ganda\\)','');
          body.replaceText ('per penurunan {{Calc Ni 2}}','');
          body.replaceText ('\\(-\\) {{Harga Ni 2}} USD/WMT','');
          }
            else {body.replaceText('{{Symbol Ni 2}}', row[columnNames.indexOf('Symbol Ni 2')]);
                  body.replaceText('{{Penalty Ni 2}}', row[columnNames.indexOf('Penalty Ni 2')]);
                      if(checkSymbol(row[columnNames.indexOf('Symbol Ni 2')]))
                      {body.replaceText('{{Pembagi Ni 2}}', 'per penurunan');}
                      else {body.replaceText('{{Pembagi Ni 2}}', 'per kenaikan');}
                  body.replaceText('{{Calc Ni 2}}', row[columnNames.indexOf('Calc Ni 2')]);
                  body.replaceText('{{Harga Ni 2}}', row[columnNames.indexOf('Harga Ni 2')]);
            };
        }
        //REJECT Ni
        if (column === 'Reject Ni') {
          if(!row[colIdx]) {
          body.replaceText('\\(Rejection {{Reject Ni}} \\)', '');
          }
            else {body.replaceText('{{Reject Ni}}', row[colIdx]);};
        }
        //CONTOH Ni        
        if (column === 'Penalty Ni 1') {
          if(!row[colIdx]) {
          body.replaceText('{{Contoh Ni}}','')
          }
            else {
                  if(!row[columnNames.indexOf('Penalty Ni 2')]) {
                  body.replaceText('{{Contoh Ni}}', 'Apabila Ni '+row[columnNames.indexOf('Symbol Ni 1')]+row[columnNames.indexOf('Penalty Ni 1')]+' = ((Aktual Ni - '+row[columnNames.indexOf('Penalty Ni 1')]+')'+'/'+row[columnNames.indexOf('Calc Ni 1')]+' x '+row[columnNames.indexOf('Harga Ni 1')]+' USD)');
                  }
                    else {body.replaceText('{{Contoh Ni}}', 'Apabila Ni '+row[columnNames.indexOf('Symbol Ni 2')]+row[columnNames.indexOf('Penalty Ni 2')]+' = (('+row[columnNames.indexOf('Penalty Ni 2')]+' - '+row[columnNames.indexOf('Penalty Ni 1')]+')/'+row[columnNames.indexOf('Calc Ni 1')]+' x '+row[columnNames.indexOf('Harga Ni 1')]+' USD) + ((Aktual Ni - '+row[columnNames.indexOf('Penalty Ni 2')]+') /'+row[columnNames.indexOf('Calc Ni 2')]+' x '+row[columnNames.indexOf('Harga Ni 2')]+' USD)');
                    }
            }
        }
        //Ni END
        
        if (column === 'Fe') {
          if (!row[colIdx]) {
            body.replaceText('Fe	:	{{Fe}}', '');
          }
            else {body.replaceText('{{Fe}}', row[colIdx]);}
          if (!row[columnNames.indexOf('Reject Fe')]) {
            body.replaceText('\\(Rejection {{Reject Fe}} \\)', '');
          }
            else {body.replaceText('{{Reject Fe}}', row[37]);}
        }
          
        if(column === 'FeNi') {
          if (!row[colIdx]) {
            body.replaceText('Fe/Ni	:	{{FeNi}}', '');
          }
            else {body.replaceText('{{FeNi}}', row[colIdx]);}
          if (!row[columnNames.indexOf('Reject FeNi')]) {
            body.replaceText('\\(Rejection {{Reject FeNi}} \\)', '');
          }
            else {body.replaceText('{{Reject FeNi}}', row[columnNames.indexOf('Reject FeNi')]);}
        }
          
        if(column === 'MC') {
          if (!row[colIdx]) {
          body.replaceText('MC	:	{{MC}}', '');
          }
          else {body.replaceText('{{MC}}', row[colIdx]);}
          //PENALTY MC
            if (!row[columnNames.indexOf('Penalty MC')]) {
              body.replaceText ('MC {{Penalty MC}}','');
              body.replaceText ('{{Harga MC}}','');
            }
              else {
                body.replaceText('{{Penalty MC}}', row[columnNames.indexOf('Penalty MC')]);
                body.replaceText('{{Harga MC}}', row[columnNames.indexOf('Harga MC')]);
              }
        }

        //SM START
        if(column === 'SM') {
          if (!row[colIdx]) {
          body.replaceText('S/M	:	{{SM}}	', '')
          }
          else {
            body.replaceText('{{SM}}', row[colIdx]);
          }
            //BONUS SM 1       
            if (!row[columnNames.indexOf('Bonus SM 1')]) {
              body.replaceText ('S/M {{Symbol Bonus SM 1}}{{Bonus SM 1}}','');
              body.replaceText ('{{Pembagi Bonus SM 1}} {{Calc Bonus SM 1}}','');
              body.replaceText ('\\(\\+\\) {{Harga Bonus SM 1}} USD/WMT','');
            }
            else {
                body.replaceText('{{Symbol Bonus SM 1}}', row[columnNames.indexOf('Symbol Bonus SM 1')]);
                body.replaceText('{{Bonus SM 1}}', row[columnNames.indexOf('Bonus SM 1')]);
                  if(!row[columnNames.indexOf('Calc Bonus SM 1')]) {
                    body.replaceText('{{Pembagi Bonus SM 1}}', '-');
                  }
                    else if(checkSymbol(row[columnNames.indexOf('Symbol Bonus SM 1')])) {
                      body.replaceText('{{Pembagi Bonus SM 1}}', 'per penurunan');}
                      {body.replaceText('{{Pembagi Bonus SM 1}}', 'per kenaikan');
                    }
                body.replaceText('{{Calc Bonus SM 1}}', row[columnNames.indexOf('Calc Bonus SM 1')]);
                body.replaceText('{{Harga Bonus SM 1}}', row[columnNames.indexOf('Harga Bonus SM 1')]);
            }
            //BONUS SM 2
           if (!row[columnNames.indexOf('Bonus SM 2')]) {
              body.replaceText ('S/M {{Symbol Bonus SM 2}}{{Bonus SM 2}} \\(Bonus ganda\\)','');
              body.replaceText ('{{Pembagi Bonus SM 2}} {{Calc Bonus SM 2}}','');
              body.replaceText ('\\(\\+\\) {{Harga Bonus SM 2}} USD/WMT','');
            }
            else {
                body.replaceText('{{Symbol Bonus SM 2}}', row[columnNames.indexOf('Symbol Bonus SM 2')]);
                body.replaceText('{{Bonus SM 2}}', row[columnNames.indexOf('Bonus SM 2')]);
                  if(!row[columnNames.indexOf('Calc Bonus SM 2')]) {
                    body.replaceText('{{Pembagi Bonus SM 2}}', '-');
                  }
                    else if(checkSymbol(row[columnNames.indexOf('Symbol Bonus SM 2')])) {
                      body.replaceText('{{Pembagi Bonus SM 2}}', 'per penurunan');}
                      {body.replaceText('{{Pembagi Bonus SM 2}}', 'per kenaikan');
                    }
                body.replaceText('{{Calc Bonus SM 2}}', row[columnNames.indexOf('Calc Bonus SM 2')]);
                body.replaceText('{{Harga Bonus SM 2}}', row[columnNames.indexOf('Harga Bonus SM 2')]);
            }

            //PENALTY SM 1
            if (!row[columnNames.indexOf('Penalty SM 1')]) {
              body.replaceText ('S/M {{Symbol SM 1}}{{Penalty SM 1}} ','');
              body.replaceText ('{{Pembagi SM 1}} {{Calc SM 1}}','');
              body.replaceText ('\\(-\\) {{Harga SM 1}} USD/WMT','');
            }
              else {
                body.replaceText('{{Symbol SM 1}}', row[columnNames.indexOf('Symbol SM 1')]);
                body.replaceText('{{Penalty SM 1}}', row[columnNames.indexOf('Penalty SM 1')]);
                  if(!row[columnNames.indexOf('Calc SM 1')]) {
                    body.replaceText('{{Pembagi SM 1}}', '-');
                  }
                    else if(checkSymbol(row[columnNames.indexOf('Symbol SM 1')])) {
                      body.replaceText('{{Pembagi SM 1}}', 'per penurunan');}
                      {body.replaceText('{{Pembagi SM 1}}', 'per kenaikan');
                    }
                body.replaceText('{{Calc SM 1}}', row[columnNames.indexOf('Calc SM 1')]);
                body.replaceText('{{Harga SM 1}}', row[columnNames.indexOf('Harga SM 1')]);
            }

            //PENALTY SM 2
            if (!row[columnNames.indexOf('Penalty SM 2')]) {
              body.replaceText ('S/M {{Symbol SM 2}}{{Penalty SM 2}} \\(Penalti ganda\\)','');
              body.replaceText ('{{Pembagi SM 2}} {{Calc SM 2}}','');
              body.replaceText ('\\(-\\) {{Harga SM 2}} USD/WMT','');
            }
              else {
                body.replaceText('{{Symbol SM 2}}', row[columnNames.indexOf('Symbol SM 2')]);
                body.replaceText('{{Penalty SM 2}}', row[columnNames.indexOf('Penalty SM 2')]);
                  if(!row[columnNames.indexOf('Calc SM 2')]) {
                    body.replaceText('{{Pembagi SM 2}}', '-');
                  }
                    else if(checkSymbol(row[columnNames.indexOf('Symbol SM 2')])) {
                      body.replaceText('{{Pembagi SM 2}}', 'per penurunan');}
                      {body.replaceText('{{Pembagi SM 2}}', 'per kenaikan');
                    }
                body.replaceText('{{Calc SM 2}}', row[columnNames.indexOf('Calc SM 2')]);
                body.replaceText('{{Harga SM 2}}', row[columnNames.indexOf('Harga SM 2')]);
            }

            //REJECT SM
            if (!row[columnNames.indexOf('Reject SM')]) {
              body.replaceText('\\(Rejection {{Reject SM}} \\)', '')
            }
              else {
                body.replaceText('{{Reject SM}}', row[columnNames.indexOf('Reject SM')])
            }
              
            //CONTOH SM
            if(!row[columnNames.indexOf('Penalty SM 1')]) {
              body.replaceText('{{Contoh SM}}','')
            }
              else {
                if(!row[columnNames.indexOf('Penalty SM 2')]) {
                body.replaceText('{{Contoh SM}}', 'Apabila SM '+row[columnNames.indexOf('Symbol SM 1')]+row[columnNames.indexOf('Penalty SM 1')]+' = (('+row[columnNames.indexOf('Penalty SM 1')]+' - Aktual SM)'+'/'+row[columnNames.indexOf('Calc SM 1')]+' x '+row[columnNames.indexOf('Harga SM 1')]+' USD)');
                }
                else {
                  body.replaceText('{{Contoh SM}}', 'Apabila SM '+row[columnNames.indexOf('Symbol SM 2')]+row[columnNames.indexOf('Penalty SM 2')]+' = (('+row[columnNames.indexOf('Penalty SM 1')]+' - '+row[columnNames.indexOf('Penalty SM 2')]+')/'+row[columnNames.indexOf('Calc SM 1')]+' x '+row[columnNames.indexOf('Harga SM 1')]+' USD) + (('+row[columnNames.indexOf('Penalty SM 2')]+' - Aktual SM) /'+row[columnNames.indexOf('Calc SM 2')]+' x '+row[columnNames.indexOf('Harga SM 2')]+' USD)');
                }
            }
          
        }
        //SM END

        //MgO
        if (column === 'MgO') {
          if (!row[colIdx]) {
          body.replaceText('MgO	:	{{MgO}}', '')
          }
          else {
            body.replaceText('{{MgO}}', row[colIdx])
          }
        }
        //PENALTY MgO
        if (column === 'Penalty MgO') {
          if (!row[colIdx]) {
          body.replaceText('MgO {{Symbol MgO}}{{Penalty MgO}}','');
          body.replaceText('{{Pembagi MgO}} {{Calc MgO}}','');
          body.replaceText('\\(-\\) {{Harga MgO}} USD/WMT','');
          }
          else {
            body.replaceText('{{Symbol MgO}}', row[columnNames.indexOf('Symbol MgO')]);
            body.replaceText('{{Penalty MgO}}', row[colIdx]);
              if(checkSymbol(row[columnNames.indexOf('Symbol MgO')])) {
                body.replaceText('{{Pembagi MgO}}', 'per penurunan');
              }
              else {
                body.replaceText('{{Pembagi MgO}}', 'per kenaikan');
              }
            body.replaceText('{{Calc MgO}}', row[columnNames.indexOf('Calc MgO')]);
            body.replaceText('{{Harga MgO}}', row[columnNames.indexOf('Harga MgO')]);
          }
        }
        if (column === 'Reject MgO')  {
          if (!row[colIdx]) {
            body.replaceText('\\(Rejection {{Reject MgO}} \\)', '')
          }
          else {
            body.replaceText('{{Reject MgO}}', row[colIdx])
          }
        }
        
        //Al2O3
        if (column === 'Al2O3') {
          if (!row[colIdx]) {
          body.replaceText('Al2O3	:	{{Al2O3}}', '')
          }
          else {
            body.replaceText('{{Al2O3}}', row[colIdx])
          }
        }
        //PENALTY Al2O3
        if (column === 'Penalty Al2O3') {
          if (!row[colIdx]) {
          body.replaceText ('Al2O3 {{Symbol Al2O3}}{{Penalty Al2O3}}','');
          body.replaceText ('{{Pembagi Al2O3}} {{Calc Al2O3}}','');
          body.replaceText ('\\(-\\) {{Harga Al2O3}} USD/WMT','');
          }
          else {
            body.replaceText('{{Symbol Al2O3}}', row[columnNames.indexOf('Symbol Al2O3')]);
            body.replaceText('{{Penalty Al2O3}}', row[colIdx]);
              if(checkSymbol(row[columnNames.indexOf('Symbol Al2O3')])) {
                body.replaceText('{{Pembagi Al2O3}}', 'per penurunan');
              }
              else {
                body.replaceText('{{Pembagi Al2O3}}', 'per kenaikan');
              }
            body.replaceText('{{Calc Al2O3}}', row[columnNames.indexOf('Calc Al2O3')]);
            body.replaceText('{{Harga Al2O3}}', row[columnNames.indexOf('Harga Al2O3')]);
          }
        }
        if (column === 'Reject Al2O3')  {
          if (!row[colIdx]) {
            body.replaceText('\\(Rejection {{Reject Al2O3}} \\)', '')
          }
          else {
            body.replaceText('{{Reject Al2O3}}', row[colIdx])
          }
        }

        //Co
        if (column === 'Co') {
          if (!row[colIdx]) {
          body.replaceText('Co	:	{{Co}}', '')
          }
          else {
            body.replaceText('{{Co}}', row[colIdx])
          }
        }
        //BONUS Co
        if(column === 'Bonus Co') {
          if(!row[colIdx]) {
            body.replaceText ('Co {{Symbol Bonus Co}}{{Bonus Co}}','');
            body.replaceText ('{{Pembagi Bonus Co}} {{Calc Bonus Co}}','');
              body.replaceText ('\\(\\+\\) {{Harga Bonus Co}} USD/WMT','');
          }
            else {
              body.replaceText('{{Symbol Bonus Co}}', row[75]);
              body.replaceText('{{Bonus Co}}', row[76]);
                if(checkSymbol(row[columnNames.indexOf('Symbol Bonus Co')])) {
                  body.replaceText('{{Pembagi Bonus Co}}', 'per penurunan');
                }
                  else {
                    body.replaceText('{{Pembagi Bonus Co}}', 'per kenaikan');
                  }
              body.replaceText('{{Calc Bonus Co}}', row[77]);
              body.replaceText('{{Harga Bonus Co}}', row[78]);
            }
        }


        //PENALTY Co
        if (column === 'Penalty Co') {
          if (!row[colIdx]) {
          body.replaceText ('Co {{Symbol Co}}{{Penalty Co}}','');
          body.replaceText ('{{Pembagi Co}} {{Calc Co}}','');
          body.replaceText ('\\(-\\) {{Harga Co}} USD/WMT','');
          }
          else {
            body.replaceText('{{Symbol Co}}', row[columnNames.indexOf('Symbol Co')]);
            body.replaceText('{{Penalty Co}}', row[colIdx]);
              if(checkSymbol(row[columnNames.indexOf('Symbol Co')])) {
                body.replaceText('{{Pembagi Co}}', 'per penurunan');
              }
              else {
                body.replaceText('{{Pembagi Co}}', 'per kenaikan');
              }
            body.replaceText('{{Calc Co}}', row[columnNames.indexOf('Calc Co')]);
            body.replaceText('{{Harga Co}}', row[columnNames.indexOf('Harga Co')]);
          }
        }
        
        if (column === 'Reject Co')  {
          if (!row[colIdx]) {
            body.replaceText('\\(Rejection {{Reject Co}} \\)', '')
          }
          else {
            body.replaceText('{{Reject Co}}', row[colIdx])
          }
        }

            
            
        //SiO2
        if (column === 'SiO2') {
          if (!row[colIdx]) {
          body.replaceText('SiO2	:	{{SiO2}}', '')
          }
          else {
            body.replaceText('{{SiO2}}', row[colIdx])
          }
        }

        //PENALTY SiO2
        if (column === 'Penalty SiO2') {
          if (!row[colIdx]) {
          body.replaceText ('SiO2 {{Symbol SiO2}}{{Penalty SiO2}}','');
          body.replaceText ('{{Pembagi SiO2}} {{Calc SiO2}}','');
          body.replaceText ('\\(-\\) {{Harga SiO2}} USD/WMT','');
          }
          else {
            body.replaceText('{{Symbol SiO2}}', row[columnNames.indexOf('Symbol SiO2')]);
            body.replaceText('{{Penalty SiO2}}', row[colIdx]);
              if(checkSymbol(row[columnNames.indexOf('Symbol SiO2')])) {
                body.replaceText('{{Pembagi SiO2}}', 'per penurunan');
              }
              else {
                body.replaceText('{{Pembagi SiO2}}', 'per kenaikan');
              }
            body.replaceText('{{Calc SiO2}}', row[columnNames.indexOf('Calc SiO2')]);
            body.replaceText('{{Harga SiO2}}', row[columnNames.indexOf('Harga SiO2')]);
          }
        }

        if (column === 'Reject SiO2')  {
          if (!row[colIdx]) {
            body.replaceText('\\(Rejection {{Reject SiO2}} \\)', '')
          }
          else {
            body.replaceText('{{Reject SiO2}}', row[colIdx])
          }
        }

        if(column === 'Size')  {
          if (!row[colIdx]) {
            body.replaceText('Size	:	{{Size}}', '')
          }
          else {
            body.replaceText('{{Size}}', row[colIdx])
          }
        }

        //PENALTY SIZE
        if(column === 'Penalty Size')  {
          if (!row[colIdx]) {
            body.replaceText ('Size {{Penalty Size}}','');
            body.replaceText ('{{Harga Size}}','');
          }
            else {
              body.replaceText('{{Penalty Size}}', row[colIdx]);
              body.replaceText('{{Harga Size}}', row[columnNames.indexOf('Harga Size')]);
            }
        }
        
        if (column === 'P') {
          if (!row[colIdx]) {
            body.replaceText('P	:	{{P}}', '')
          }
            else {
              body.replaceText('{{P}}', row[colIdx])
            }
        }

        if(column === 'Reject P') {
          if (!row[colIdx]) {
            body.replaceText('\\(Rejection {{Reject P}} \\)', '')
          }
          else {
            body.replaceText('{{Reject P}}', row[colIdx])
          }
        }
            
        
         
        if (column === 'S') {
          if (!row[colIdx]) {
            body.replaceText('S	:	{{S}}', '')
          }
            else {
              body.replaceText('{{S}}', row[colIdx])
            }
        }

        if(column === 'Reject S') {
          if (!row[colIdx]) {
            body.replaceText('\\(Rejection {{Reject S}} \\)', '')
          }
          else {
            body.replaceText('{{Reject S}}', row[colIdx])
          }
        }
        

        if (column === 'CaO') {
          if (!row[colIdx]) {
            body.replaceText('CaO	:	{{CaO}}', '')
          }
            else {
              body.replaceText('{{CaO}}', row[colIdx])
            }
        }

        if(column === 'Reject CaO') {
          if (!row[colIdx]) {
            body.replaceText('\\(Rejection {{Reject CaO}} \\)', '')
          }
          else {
            body.replaceText('{{Reject CaO}}', row[colIdx])
          }
        }


        //BONUS QTY
        if(column === 'Bonus Qty 1') {
          body.replaceText('{{Bonus Qty 1}}', row[colIdx]);
          if(!row[colIdx]) {
            body.replaceText('{{Harga Bonus Qty 1}}', '');
          }
            else {
              body.replaceText('{{Harga Bonus Qty 1}}', '\(+\) '+row[columnNames.indexOf('Harga Bonus Qty 1')]+' USD/WMT');
            }
        }

        if(column === 'Bonus Qty 2') {
        body.replaceText('{{Bonus Qty 2}}', row[colIdx]);
          if(!row[colIdx]) {
            body.replaceText('{{Harga Bonus Qty 2}}', '');
          }
            else {
              body.replaceText('{{Harga Bonus Qty 2}}', '\(+\) '+row[columnNames.indexOf('Harga Bonus Qty 2')]+' USD/WMT');
            }
        }

        if(column === 'Bonus Qty 3') {
        body.replaceText('{{Bonus Qty 3}}', row[colIdx]);
          if(!row[colIdx]) {
            body.replaceText('{{Harga Bonus Qty 3}}', '');
          }
            else {
              body.replaceText('{{Harga Bonus Qty 3}}', '\(+\) '+row[columnNames.indexOf('Harga Bonus Qty 3')]+' USD/WMT');
            }
        }
        //SPECS END


        //PEMBAYARAN
        if(column === 'Down Payment') {
          if (!row[colIdx]) {
            body.replaceText('{{Down Payment}}', '');
            body.replaceText('Pembayaran Awal', '');
          }
          else {
            body.replaceText('{{Down Payment}}', row[colIdx]);
          }
        }

        if(column === 'Provisional Payment') {
          if (!row[colIdx]) {
            body.replaceText('{{Provisional Payment}}', '');
            body.replaceText('Pembayaran Provisional', '');
          }
          else {
            body.replaceText('{{Provisional Payment}}', row[colIdx]);
          }
        }


        body.replaceText(`{{${column}}}`, row[colIdx])
      })

          
      

      body.replaceText('`','');

      doc.saveAndClose();
      const url = doc.getUrl();
      Logger.log(columnNames)
      Logger.log(columnNames.indexOf)
      sheet.getRange(index + 1, columnNames.indexOf('Document Link')+1).setValue(url)
})
}

