import { Component } from "@angular/core";
import { RouterOutlet } from "@angular/router";
import * as ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { MatDatepickerModule } from "@angular/material/datepicker";
import { MatNativeDateModule } from "@angular/material/core";
import { MatFormFieldModule } from "@angular/material/form-field";
import { MatInputModule } from "@angular/material/input";
import { CommonModule } from "@angular/common";
import { ReactiveFormsModule, FormControl } from "@angular/forms";
import { Console } from "node:console";
@Component({
  selector: "app-root",
  standalone: true,
  imports: [
    RouterOutlet,
    CommonModule,
    ReactiveFormsModule,
    MatDatepickerModule,
    MatNativeDateModule,
    MatFormFieldModule,
    MatInputModule,
  ],
  template: `
    <div>
      <h1>Jelenléti ív készítés</h1>
      <label for="nameInput">Név:</label>
      <input id="nameInput" type="text" name="nev" />
      <br />
      <br />
      <p>Év, Hónap:</p>
      <mat-form-field appearance="fill">
        <mat-label *ngIf="dateStr">{{ dateStr }}</mat-label>
        <input
          matInput
          [matDatepicker]="picker"
          placeholder="Select month and year"
          (focus)="picker.open()"
          readonly="True"
        />
        <mat-datepicker
          #picker
          startView="multi-year"
          (monthSelected)="chosenMonthHandler($event, picker)"
          panelClass="month-picker"
        >
        </mat-datepicker>
      </mat-form-field>
      <br />
      <br />
      <label for="dayOff">Szabadnapok, ünnepnapok:</label>
      <input
        id="dayOffInput"
        type="text"
        name="dayOff"
        placeholder="pl.: 15,18"
      />
      <br />
      <br />
      <label for="companySelect">Cég:</label>
      <select id="companySelect">
        <option value="Molnár-Kárpát Kft.">Molnár-Kárpát Kft.</option>
        <option value="KEGA-Kárpát Kft.">KEGA-Kárpát Kft.</option>
      </select>
      <br />
      <br />
      <label for="workHoursSelect">Óraszám:</label>
      <select id="workHoursSelect">
        <option value="4">4</option>
        <option value="8">8</option>
      </select>
      <br />
      <br />
      <button (click)="downloadPdf()">Jelenléti ív letöltése</button>
      <br />
      <br />
      <button (click)="openPdf()">Jelenléti ív megnyitása</button>
    </div>

    <router-outlet />
  `,
  styles: [
    `
      div {
        max-width: 500px;
        margin-top: 20px;
        margin-right: auto;
        margin-left: auto;
        padding: 20px;
        border: 1px solid #ccc;
        border-radius: 8px;
        background-color: #f9f9f9;
      }

      h1 {
        text-align: center;
        color: #333;
        font-family: Arial, sans-serif;
        margin-bottom: 20px;
      }

      label {
        font-family: Arial, sans-serif;
        font-size: 14px;
        color: #333;
        margin-bottom: 5px;
        display: block;
      }

      input[type="text"],
      select {
        width: 100%;
        padding: 8px;
        margin-bottom: 10px;
        border: 1px solid #ccc;
        border-radius: 4px;
        font-family: Arial, sans-serif;
        font-size: 14px;
      }

      mat-form-field {
        width: 100%;
        margin-bottom: 16px;
      }

      button {
        width: 100%;
        padding: 10px;
        background-color: #4caf50;
        color: white;
        border: none;
        border-radius: 4px;
        font-size: 16px;
        font-family: Arial, sans-serif;
        cursor: pointer;
        margin-top: 20px;
      }

      button:hover {
        background-color: #45a049;
      }

      p {
        font-family: Arial, sans-serif;
        font-size: 14px;
        color: #333;
        margin-bottom: 5px;
      }
    `,
  ],
})
export class AppComponent {
  async getBlob() {
    const nameInputValue = (
      document.getElementById("nameInput") as HTMLInputElement
    ).value;
    const dayOffInputValue = (
      document.getElementById("dayOffInput") as HTMLInputElement
    ).value;
    const companySelectValue = (
      document.getElementById("companySelect") as HTMLInputElement
    ).value;
    const workHoursSelectValue = (
      document.getElementById("workHoursSelect") as HTMLInputElement
    ).value;
    console.log(
      nameInputValue,
      dayOffInputValue,
      this.selectedDate,
      companySelectValue,
      workHoursSelectValue
    );

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("nameInputValue");

    // Fejléc
    worksheet.mergeCells("A1:F1");
    const titleCell = worksheet.getCell("A1");
    titleCell.value = `Jelenléti ív`;
    titleCell.font = { bold: true, size: 14 };

    //Cégnév Nagy
    worksheet.mergeCells("A2:C3");
    const companyCell = worksheet.getCell("A2");
    companyCell.value = companySelectValue;
    companyCell.font = { bold: true, size: 20 };

    //Dolgozó Név
    worksheet.mergeCells("D2:F2");
    const nameCell = worksheet.getCell("D2");
    nameCell.value = "Név: " + nameInputValue;
    nameCell.font = { bold: true };

    //Cégnév Kicsi
    worksheet.mergeCells("D3:F3");
    const compNameCell = worksheet.getCell("D3");
    compNameCell.value = "Cég: " + companySelectValue;
    compNameCell.font = { bold: true };

    //ÉvHónap
    worksheet.mergeCells("A4:B4");
    const yearMonthCell = worksheet.getCell("A4");
    yearMonthCell.value = this.dateStr;
    yearMonthCell.font = { bold: true };

    //FELIRATOK:
    //Érkezés
    const arrivalCell = worksheet.getCell("C4");
    arrivalCell.value = "Érkezés";
    arrivalCell.font = { bold: true };
    //Távozás
    const exitCell = worksheet.getCell("D4");
    exitCell.value = "Távozás";
    exitCell.font = { bold: true };
    //Óraszám
    const workHoursCell = worksheet.getCell("E4");
    workHoursCell.value = "Óraszám";
    workHoursCell.font = { bold: true };
    //Aláírás
    const signCell = worksheet.getCell("F4");
    signCell.value = "Aláírás";
    signCell.font = { bold: true };

    const daysHu = [
      "Hétfő",
      "Kedd",
      "Szerda",
      "Csütörtök",
      "Péntek",
      "Szombat",
      "Vasárnap",
    ];
    const firstDayIndex = daysHu.findIndex(
      (day) =>
        day.toLocaleLowerCase() === this.getFirstDayName(this.selectedDate)
    );
    console.log(firstDayIndex);

    let monthDays = new Date(
      this.selectedDate.getFullYear(),
      this.selectedDate.getMonth() + 1,
      0
    ).getDate();

    for (let i = 1; i <= monthDays; i++) {
      var row = i + 4;
      var rowSt = row.toString();
      //Napok sorszáma
      worksheet.getCell("A" + rowSt).value = i;

      //Napok neve
      const currentDayName = daysHu[(i + firstDayIndex - 1) % 7];
      worksheet.getCell("B" + rowSt).value = currentDayName;

      const c = worksheet.getCell("C" + rowSt);
      const d = worksheet.getCell("D" + rowSt);
      const e = worksheet.getCell("E" + rowSt);
      //Ünnepnapok, hétvégék (szürke)
      if (
        currentDayName === "Szombat" ||
        currentDayName === "Vasárnap" ||
        dayOffInputValue.split(",").includes(i.toString())
      ) {
        c.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "808080" },
        };
        d.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "808080" },
        };
        e.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "808080" },
        };
      } else {
        //Munkaóra feliratok
        c.value = "8:00";
        d.value = workHoursSelectValue === "4" ? "12:00" : "16:30";
        e.value = workHoursSelectValue;
      }

      //Középre igazítás
      for (let row = 1; row <= monthDays + 4; row++) {
        for (let col = 1; col <= 6; col++) {
          var cell = worksheet.getCell(row, col);
          cell.alignment = {
            horizontal: "center",
            vertical: "middle",
          };
          cell.border = {
            top: { style: "thin" },
            left: { style: "thin" },
            bottom: { style: "thin" },
            right: { style: "thin" },
          };
        }
      }
    }

    //Oszlopszélesség
    for (let i = 1; i < 6; i++) {
      worksheet.getColumn(i).width = 12;
    }
    worksheet.getColumn(6).width = 24;

    
    //Xlsx letöltés
    const buffer = await workbook.xlsx.writeBuffer();
    return new Blob([buffer], { type: "application/octet-stream" });
  }

  async downloadPdf()
  {
    const blob = this.getBlob();
    const nameInputValue = (
      document.getElementById("nameInput") as HTMLInputElement
    ).value;
    saveAs(await blob, nameInputValue + " jelenléti "+this.dateStr.split(',')[1].trim()+".xlsx");
  }

  async openPdf()
  {
    const blob = this.getBlob();
    var fileURL = URL.createObjectURL(await blob);
    window.open(fileURL, '_blank');
  }

  dateStr: string = "";
  selectedDate: Date = new Date();

  chosenMonthHandler(normalizedMonth: Date, datepicker: any) {
    console.log("Honap kivalasztva");
    this.dateStr = `${normalizedMonth.getFullYear()}, ${normalizedMonth.toLocaleString(
      "default",
      { month: "long" }
    )}`;
    this.selectedDate = normalizedMonth;
    datepicker.close();
  }

  getFirstDayName(date: Date) {
    var a = date.toLocaleDateString("hu-HU", { weekday: "long" });
    console.log(a);
    return a;
  }
}
