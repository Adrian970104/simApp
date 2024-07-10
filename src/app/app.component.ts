import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import * as ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet],
  template: `
    <div>
      <h1>Jelenléti ív készítés</h1>
      <label for="nameInput">Név:</label>
      <input id="nameInput" type="text" name="nev"/>
      <br/>
      <label for="month">Év, Hónap:</label>
      <input id="monthInput" type="text" name="month"/>
      <br/>
      <label for="monthDayCount">Hónap napjai:</label>
      <input id="monthDayCountInput" type="text" name="monthDayCount"/>
      <br/>
      <label for="dayOff">Szabadnapok, ünnepnapok:</label>
      <input id="dayOffInput" type="text" name="dayOff"/>
      <br/>
      <label for="firstDay">Elsőnap:</label>
      <input id="firstDayInput" type="text" name="firstDay"/>
      <br/>
      <label for="companySelect">Cég:</label>
      <select id="companySelect">
        <option value="Molnár-Kárpát Kft.">Molnár-Kárpát Kft.</option>
        <option value="KEGA-Kárpát Kft.">KEGA-Kárpát Kft.</option>
      </select>
      <br />
      <label for="workHoursSelect">Óraszám:</label>
      <select id="workHoursSelect">
        <option value="4">4</option>
        <option value="8">8</option>
      </select>
      <br />
      <button (click)="logName()">Log Name</button>
      <br />
      <button (click)="generateXlsx()">Jelenléti Letöltése!</button>
    </div>

    <router-outlet />
  `,
  styles: [],
})
export class AppComponent {
  title = 'simApp';
  nev: string = '';
  month: string = '';
  dayOff: string = '';
  firstDay: string = '';
  company: string = '';
  workHours: string = '';

  logName() {
    const nameInputValue = (document.getElementById('nameInput') as HTMLInputElement).value;
    console.log('Entered name:', nameInputValue);
  }

  async generateXlsx() {
    const nameInputValue = (document.getElementById('nameInput') as HTMLInputElement).value;
    const monthInputValue = (document.getElementById('monthInput') as HTMLInputElement).value;
    const monthDayCountInputValue = (document.getElementById('monthDayCountInput') as HTMLInputElement).value;
    const dayOffInputValue = (document.getElementById('dayOffInput') as HTMLInputElement).value;
    const firstDayInputValue = (document.getElementById('firstDayInput') as HTMLInputElement).value;
    const companySelectValue = (document.getElementById('companySelect') as HTMLInputElement).value;
    const workHoursSelectValue = (document.getElementById('workHoursSelect') as HTMLInputElement).value;
    console.log(nameInputValue, monthInputValue, dayOffInputValue, firstDayInputValue, companySelectValue, workHoursSelectValue);

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Sheet1');

    // Fejléc
    worksheet.mergeCells('A1:F1');
    const titleCell = worksheet.getCell('A1');
    titleCell.value = `Jelenléti ív`;
    titleCell.font = { bold: true, size: 14 };

    //Cégnév Nagy
    worksheet.mergeCells('A2:C3');
    const companyCell = worksheet.getCell('A2');
    companyCell.value = companySelectValue;
    companyCell.font = { bold: true, size: 20 };

    //Dolgozó Név
    worksheet.mergeCells('D2:F2');
    const name = worksheet.getCell('D2');
    name.value = "Név: " + nameInputValue;
    name.font = { bold: true };

    //Cégnév Kicsi
    worksheet.mergeCells('D3:F3');
    const compName = worksheet.getCell('D3');
    compName.value = companySelectValue;
    compName.font = { bold: true };

    //ÉvHónap
    worksheet.mergeCells('A4:B4');
    const mon = worksheet.getCell('A4');
    mon.value = monthInputValue;
    mon.font = { bold: true };

    //FELIRATOK:
    //Érkezés
    const erk = worksheet.getCell('C4');
    erk.value = "Érkezés";
    erk.font = { bold: true };
    //Távozás
    const tav = worksheet.getCell('D4');
    tav.value = "Távozás";
    tav.font = { bold: true };
    //Óraszám
    const ora = worksheet.getCell('E4');
    ora.value = "Óraszám";
    ora.font = { bold: true };
    //Aláírás
    const alair = worksheet.getCell('F4');
    alair.value = "Aláírás";
    alair.font = { bold: true };

    const days = ["Hétfő", "Kedd", "Szerda", "Csütörtök", "Péntek", "Szombat", "Vasárnap"];
    const firstDayIndex = days.indexOf(firstDayInputValue);
    if (firstDayIndex < 0)
    {
      alert("Invalid day name. Can be: Hétfő, Kedd, Szerda, Csütörtök, Péntek, Szombat, Vasárnap");
      throw new Error("Invalid day name. Can be: Hétfő, Kedd, Szerda, Csütörtök, Péntek, Szombat, Vasárnap");
    }

    for (let i = 1; i <= parseInt(monthDayCountInputValue); i++) {
      var row = i + 4;
      var rowSt = row.toString();
      worksheet.getCell('A' + rowSt).value = i;

      const currentDayName = days[(i + firstDayIndex - 1) % 7];
      worksheet.getCell('B' + rowSt).value = currentDayName;
      
      const c = worksheet.getCell('C' + rowSt);
      const d = worksheet.getCell('D' + rowSt);
      const e = worksheet.getCell('E' + rowSt);
      if (currentDayName === "Szombat" || currentDayName === "Vasárnap" || dayOffInputValue.split(',').includes(i.toString()))
      {
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '808080' } };
        d.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '808080' } };
        e.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '808080' } };
      } else {
        c.value = "8:00";
        d.value = workHoursSelectValue === '4' ? "12:00" : "16:30";
        e.value = workHoursSelectValue;
      }
      
      for (let row = 1; row <= parseInt(monthDayCountInputValue) + 4; row++) {
        for (let col = 1; col <= 6; col++) {
          var cell = worksheet.getCell(row, col);
          cell.alignment = {
            horizontal : 'center',
            vertical : 'middle'
          }
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        }
      }
    }

    for(let i = 1; i<6; i++)
    {
      worksheet.getColumn(i).width = 12;
    }
    worksheet.getColumn(6).width = 24;
    
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: 'application/octet-stream' });
    saveAs(blob, nameInputValue + 'Jelenléti.xlsx');
  }

}
