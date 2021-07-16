// @ts-nocheck
import { Component, ViewEncapsulation, ViewChild } from '@angular/core';
import { getUSEntityData, getNonUSEntityData, getBusinessLicensesData } from './data';
import { Spreadsheet, SheetModel, CellRenderEventArgs, getRangeIndexes, getCellAddress, CellSaveEventArgs } from '@syncfusion/ej2-spreadsheet';
import { CellModel, BeforeSelectEventArgs, getRow, getRangeAddress, getColumnsWidth, getSwapRange, getCell } from '@syncfusion/ej2-spreadsheet';
import { RowModel, isNumber, getColumn, MenuSelectEventArgs } from '@syncfusion/ej2-spreadsheet';
import { HeaderModel, SelectEventArgs, BeforeOpenCloseMenuEventArgs, MenuItemModel } from '@syncfusion/ej2-navigations';
import { closest, EventHandler, getComponent, select, selectAll, detach } from '@syncfusion/ej2-base';
import { ButtonModel } from '@syncfusion/ej2-buttons';
import { DataManager, Query } from '@syncfusion/ej2-data';
import { DropDownButton, ItemModel, MenuEventArgs } from '@syncfusion/ej2-splitbuttons';
import { ChangeEventArgs } from '@syncfusion/ej2-dropdowns';
import { Dialog } from '@syncfusion/ej2-popups';
import { Popup } from '@syncfusion/ej2-angular-popups';
import { createCheckBox } from '@syncfusion/ej2-angular-buttons';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css'],
  encapsulation: ViewEncapsulation.None
})
export class AppComponent {
  @ViewChild('default')
  public spreadsheetObj: Spreadsheet;
  @ViewChild('bulkActionMenu')
  public bulkActionMenu: DropDownButton;
  @ViewChild('bulkEditDialog')
  public bulkEditDlgObj: Dialog;
  public usEntityData: Object[] = getUSEntityData();
  public nonUSEntityData: Object[] = getNonUSEntityData();
  public businessLicensesData: Object[] = getBusinessLicensesData();
  public tabHeaders: HeaderModel[] = [
    { text: 'US Entities' }, { text: 'Non-US Entities' }, { text: 'Bussiness Licenses' }
  ];
  public entityNameData: Object[] = [
    { Name: 'Mountains Heart LLC' }, { Name: 'John & Brothers LLC' }, { Name: 'Free Mountains LLC' },
    { Name: 'Albert Sun LLC' }, { Name: 'Albert & Co LLC' }
  ];
  public fields: Object = { value: 'Name' };
  public openUrl = 'https://ej2services.syncfusion.com/production/web-services/api/spreadsheet/open';
  public saveUrl = 'https://ej2services.syncfusion.com/production/web-services/api/spreadsheet/save';
  private isCreated: boolean;
  private autoFillAddress: number[];
  private curAddress: string = '';
  public bulkActionMenuItems: ItemModel[] = [
    { text: 'Bulk Edit' }, { text: 'Export' }, { text: 'Remove from Quote' }
  ];
  public columnData: Object[] = [];
  public columnValue: string = '';
  public columnFields: Object = { value: 'Name' };
  public buttons: { [key: string]: ButtonModel }[] = [
    { click: this.bulkEditSave.bind(this), buttonModel: { content: 'Sure', isPrimary: true } },
    { click: this.bulkEditDlgClose.bind(this), buttonModel: { content: 'Cancel' } }
  ];

  beforeLoad(): void {
    if (this.isCreated) { return; }
    this.spreadsheetObj.sheets.forEach((sheet: SheetModel): void => {
      this.spreadsheetObj.cellFormat({ fontWeight: 'bold', textAlign: 'center', verticalAlign: 'middle' }, sheet.name + '!A1:F1');
    });
    this.spreadsheetObj.cellFormat({ textAlign: 'center', verticalAlign: 'middle' }, 'US Entities!A2:A100');
    this.spreadsheetObj.cellFormat({ textAlign: 'center', verticalAlign: 'middle' }, 'US Entities!C2:F100');
    this.spreadsheetObj.cellFormat({ textAlign: 'right', verticalAlign: 'middle' }, 'US Entities!G2:G100');
    this.spreadsheetObj.numberFormat('mm-dd-yyyy', 'Non-US Entities!E2:E100');
  }

  created() {
    this.spreadsheetObj.applyFilter(null, 'B1:F100');
    this.bulkActionMenu.element.disabled = !this.spreadsheetObj.getActiveSheet().selectedRange.includes(' ');
    window.addEventListener('resize', this.onResize.bind(this));
    EventHandler.add(select('.e-sheet-panel', this.spreadsheetObj.element), 'mousedown touchstart', this.autofillMouseDown, this);
    this.isCreated = true;
  }

  onAutoCompleteChange(args: ChangeEventArgs): void {
    this.spreadsheetObj.updateCell({ value: args.value }, this.spreadsheetObj.getActiveSheet().activeCell);
  }

  autofillMouseDown(e: MouseEvent): void {
    if (!(e.target as Element).classList.contains('e-selection') && !(e.target as Element).classList.contains('e-active-cell')) {
      return;
    }
    this.autoFillAddress = getRangeIndexes(this.spreadsheetObj.getActiveSheet().selectedRange);
    select('.e-sheet', this.spreadsheetObj.element).classList.add('e-autofill-start');
    EventHandler.add(document, 'mouseup touchend', this.autofillMouseUp, this);
  }

  autofillMouseUp(): void {
    select('.e-sheet', this.spreadsheetObj.element).classList.remove('e-autofill-start');
    EventHandler.remove(document, 'mouseup touchend', this.autofillMouseUp);
    this.setAutoFill();
    this.autoFillAddress = null;
    this.curAddress = null;
  }

  onResize(): void {
    this.spreadsheetObj.resize();
  }

  onContextMenuBeforeOpen(args: BeforeOpenCloseMenuEventArgs): void {
    if (args.parentItem === null) {
      const sheet: SheetModel = this.spreadsheetObj.getActiveSheet();
      const menuItems: MenuItemModel[] = [];
      const selectRange: number[] = getRangeIndexes(sheet.selectedRange);
      if (selectRange[1] <= sheet.usedRange.colIndex) {
        menuItems.push({ text: Math.abs(selectRange[3] - selectRange[1]) > 0 ? 'Hide Columns' : 'Hide Column' });
      }
      let hiddenCount: number = 0;
      for (let i: number = 0, len: number = sheet.usedRange.colIndex; i < len; i++) {
        if (getColumn(sheet, i).hidden) { hiddenCount++; }
      }
      if (hiddenCount) {
        menuItems.push({ text: hiddenCount === 1 ? 'Show Column' : 'Show Columns' });
      }
      if (selectRange[0] > 0) {
        menuItems.push({ text: 'Insert Row' });
      }
      if (!menuItems.length) { return; }
      menuItems.push({ separator: true });
      this.spreadsheetObj.addContextMenuItems(menuItems, 'Hyperlink');
    }
  }

  onContextMenuItemSelect(args: MenuSelectEventArgs): void {
    if (args.item.text.includes('Hide')) {
      const selectRange: number[] = getRangeIndexes(this.spreadsheetObj.getActiveSheet().selectedRange);
      this.spreadsheetObj.hideColumn(selectRange[1], selectRange[3]);
    } else if (args.item.text.includes('Show')) {
      const sheet: SheetModel = this.spreadsheetObj.getActiveSheet();
      this.spreadsheetObj.hideColumn(0, sheet.usedRange.colIndex, false);
      requestAnimationFrame(() => {
        const parentCells: HTMLElement[] = [].slice.call(selectAll( '.e-cell .e-drop-icon', this.spreadsheetObj.element ));
        if (parentCells.length) {
          let cell: HTMLElement = closest(parentCells[parentCells.length - 1], '.e-cell');
          if (cell) {
            let rowIdx: number = Number(cell.parentElement.getAttribute('aria-rowindex')) - 1;
            this.spreadsheetObj.updateCell({}, getCellAddress(rowIdx, 1));
            this.onCellRender({ cell: getCell(rowIdx, 1, sheet, false, true), element: cell, address: getCellAddress(rowIdx, 1),
              rowIndex: rowIdx, colIndex: 1 });
          }
        }
      });
    } else if (args.item.text === 'Insert Row') {
      const sheet: SheetModel = this.spreadsheetObj.getActiveSheet();
      const selectRange: number[] = getRangeIndexes(sheet.selectedRange);
      if (sheet.name === 'US Entities') {
        this.spreadsheetObj.insertRow([{ index: selectRange[0], height: 34 }]);
        this.spreadsheetObj.cellFormat({ textAlign: 'center', verticalAlign: 'middle' }, getCellAddress(selectRange[0], 0));
        this.spreadsheetObj.cellFormat({ textAlign: 'center', verticalAlign: 'middle' }, getCellAddress(
          selectRange[0], 2) + ':' + getCellAddress(selectRange[0], 5));
        this.spreadsheetObj.cellFormat({ textAlign: 'right', verticalAlign: 'middle' }, getCellAddress(selectRange[0], 6));
      } else {
        this.spreadsheetObj.insertRow(selectRange[0]);
        if (sheet.name === 'Non-US Entities') {
          requestAnimationFrame((): void => this.spreadsheetObj.numberFormat('mm-dd-yyyy', getCellAddress(selectRange[0], 4)));
        }
      }
    }
  }

  tabSelected(args: SelectEventArgs): void {
    this.spreadsheetObj.activeSheetIndex = args.selectedIndex;
  }

  counter(i: number) {
    return new Array(i);
  }

  queryCellInfo(args: CellRenderEventArgs): void {
    const sheet: SheetModel = this.spreadsheetObj.getActiveSheet();
    if (args.colIndex === 1 && args.cell && args.cell.value && !args.cell.colSpan && !getRow(sheet, args.rowIndex).cells[args.colIndex + 1].value) {
      args.cell.colSpan = sheet.usedRange.colIndex - 1;
      this.spreadsheetObj.cellFormat({ backgroundColor: '#deecf9' }, getCellAddress(args.rowIndex, 0) + ':' + getCellAddress(args.rowIndex, sheet.usedRange.colIndex));
    }
  }

  onCellRender(args: CellRenderEventArgs): void {
    const sheet: SheetModel = this.spreadsheetObj.getActiveSheet();
    if (sheet.name === 'US Entities' && args.rowIndex < sheet.rows.length) {
      if (args.colIndex === 0) {
        const checkbox: HTMLElement = createCheckBox(this.spreadsheetObj.createElement, false, { checked:
          this.isCreated && sheet.selectedRange.includes(args.address + ':' + getCellAddress(args.rowIndex, sheet.colCount - 1)) });
        args.element.appendChild(checkbox);
      } else if (args.rowIndex > 0) {
        if (args.colIndex === 1 && args.cell && args.cell.value && args.cell.colSpan > 1) {
          args.element.firstElementChild.style.width = getColumnsWidth(sheet, args.colIndex, args.colIndex + 1) + 'px';
          const arrowSpan: HTMLElement = this.spreadsheetObj.createElement('span', { className: 'e-drop-icon e-icons' });
          if (!getRow(sheet, args.rowIndex).cells[0].style.borderLeft) {
            args.element.classList.add('e-collapsed');
          }
          arrowSpan.addEventListener('click', this.onParentClick.bind(this));
          args.element.appendChild(arrowSpan);
        } else if (args.colIndex === 2 && args.cell && !args.cell.validation) {
          this.setListDataValidation('Delaware,Colorado,Texas', args.address);
          this.updateDependantValidation(args.cell.value || 'Delaware', args.rowIndex, args.colIndex, false);
        }
      }
    } else if ((args.colIndex === 1 || args.colIndex === 6) && args.element.firstElementChild) {
      (args.element.firstElementChild as HTMLElement).style.display = 'none';
    }
  }

  onParentClick(e: MouseEvent): void {
    let target: HTMLElement = closest(e.target as HTMLElement, '.e-cell') as HTMLElement;
    const rowIndex: number = Number(closest(target, '.e-row').getAttribute('aria-rowindex')) - 1;
    const sheet: SheetModel = this.spreadsheetObj.getActiveSheet();
    let result: { EntityChild: any[] }[] = new DataManager(sheet.ranges[0].dataSource).executeLocal(
      new Query()
        .select(['Entity/License Holder Name', 'EntityChild'])
        .where('Entity/License Holder Name', 'equal', getCell(rowIndex, 1, sheet).value)
        .where('EntityChild', 'notequal', undefined)) as { EntityChild: any[] }[];
    if (!result || !result.length || !result[0].EntityChild) {
      return;
    }
    const childData: any[] = result[0].EntityChild;
    if (target.classList.contains('e-collapsed')) {
      target.classList.remove('e-collapsed');
      const rows: RowModel[] = [];
      let cellValue: string;
      for (let i: number = 0, len: number = childData.length; i < len; i++) {
        rows.push({ cells: [], height: 34 });
        if (i === 0) {
          rows[0].index = rowIndex + 1;
        }
        for (let j: number = 0, colLen: number = sheet.usedRange.colIndex; j <= colLen; j++) {
          cellValue = getCell(0, j, sheet, false, true).value;
          rows[i].cells.push(j === 0 || !cellValue ? {} : { value: childData[i][cellValue] });
          if (j !== 1) {
            rows[i].cells[rows[i].cells.length - 1].style = { textAlign: j === colLen ? 'right' : 'center', verticalAlign: 'middle' };
            if (j === 0) {
              rows[i].cells[rows[i].cells.length - 1].style.borderLeft = '2px solid #0078d6';
            }
          }
          if (j === 2) {
            rows[i].cells[rows[i].cells.length - 1].validation = { type: 'List',  operator: 'Between', value1: 'Delaware,Colorado,Texas',
              ignoreBlank: true, inCellDropDown: true, isHighlighted: true };
          }
          if (j === 3) {
            cellValue = this.updateDependantValidation(rows[i].cells[rows[i].cells.length - 1].value, null, null, false, true);
            if (cellValue) {
              rows[i].cells[rows[i].cells.length - 1].validation = { type: 'List',  operator: 'Between', value1: cellValue, ignoreBlank: true,
                inCellDropDown: true, isHighlighted: true };
            }
          }
        }
      }
      this.spreadsheetObj.insertRow(rows);
      this.spreadsheetObj.setBorder({ borderLeft: '2px solid #0078d6' }, getCellAddress(rowIndex, 0));
    } else {
      target.classList.add('e-collapsed');
      this.spreadsheetObj.delete(rowIndex + 1, rowIndex + childData.length, 'Row');
      this.spreadsheetObj.setBorder({ borderLeft: '' }, getCellAddress(rowIndex, 0));
    }
  }

  onBtnClick(e: MouseEvent): void {
    const target: HTMLElement = e.target as HTMLElement;
    const popup: HTMLElement = select('.e-dropdown-popup.custom-popup');
    let dropDown: Popup;
    if (popup) {
      dropDown = getComponent(popup, Popup) as Popup;
      dropDown.relateTo = target;
    } else {
      const div: HTMLElement = this.spreadsheetObj.createElement('div', { className: 'e-dropdown-popup custom-popup' });
      const ul: HTMLElement = div.appendChild(this.spreadsheetObj.createElement('ul', { attrs: { 'role': 'menu', 'tabindex': '0' } }));
      const menuItems: string[] = ['Add new jurisdiction', 'Add Alert', 'Add Comment', 'Remove from Quote'];
      menuItems.forEach((item: string): void => {
        ul.appendChild(this.spreadsheetObj.createElement(
          'li', { className: 'e-item e-btn-item', innerHTML: item, attrs: { 'role': 'menuItem', 'tabindex': '-1' } }));
      });
      document.body.appendChild(div);
      dropDown = new Popup(div, {
        relateTo: target,
        collision: { X: 'fit', Y: 'flip' },
        position: { X: 'left', Y: 'bottom' },
        targetType: 'relative'
      });
      if (dropDown.element.style.position === 'fixed') {
        dropDown.refreshPosition(target);
      }
      EventHandler.add(ul, 'click', this.cellMenuClick, this);
      EventHandler.add(document, 'mousedown touchstart', this.cellMenuMouseDown, this);
    }
    dropDown.show();
  }

  cellMenuMouseDown(e: MouseEvent): void {
    const target: HTMLElement = e.target as HTMLElement;
    if (!target.classList.contains('e-btn-menu') && !target.classList.contains('e-btn-item')) {
      this.closeCellMenuPopup();
    }
  }

  closeCellMenuPopup(): void {
    const popup: HTMLElement = select('.e-dropdown-popup.custom-popup');
    if (popup && !popup.classList.contains('e-popup-close')) {
      (getComponent(popup, Popup) as Popup).hide();
    }
  }

  cellMenuClick(e: MouseEvent): void {
    this.closeCellMenuPopup();
    const text: string = (e.target as HTMLElement).textContent;
    if (text === 'Remove from Quote') {
      const rowIndex: number = getRangeIndexes(this.spreadsheetObj.getActiveSheet().selectedRange)[0];
      this.spreadsheetObj.delete(rowIndex, rowIndex, 'Row');
    }
  }

  setListDataValidation(values: string, address: string): void {
    this.spreadsheetObj.addDataValidation(
      { type: 'List',  operator: 'Between', value1: values, ignoreBlank: true, inCellDropDown: true, isHighlighted: true },
      address
    );
  }

  updateDependantValidation(jurisdiction: string, rowIdx: number, colIdx: number, updateValue: boolean, returnValue?: boolean): string {
    if (colIdx === 2) {
      let values: string = '';
      let updateValue: string = '';
      switch (jurisdiction) {
        case 'Delaware':
          values = 'LLC,Corp';
          updateValue = 'LLC';
          break;
        case 'Colorado':
          values = 'ColoradoType1,ColoradoType2';
          updateValue = 'ColoradoType1';
          break;
        case 'Texas':
          values = 'TexasType1,TexasType2';
          updateValue = 'TexasType1';
          break;
      }
      if (returnValue || !values) {
        return values;
      }
      const dependantAddress: string = getCellAddress(rowIdx, colIdx + 1);
      if (updateValue) {
        this.spreadsheetObj.updateCell({ value: updateValue }, dependantAddress);
      }
      this.setListDataValidation(values, dependantAddress);
    }
    return '';
  }

  onCellSave(args: CellSaveEventArgs) {
    const indexes: number[] = getRangeIndexes(args.address);
    this.updateDependantValidation(args.value, indexes[0], indexes[1], true);
  }

  onBeforeSelect(args: BeforeSelectEventArgs): void {
    if (!this.isCreated) { return; }
    const rangeIndexes: number[] = getRangeIndexes(args.range);
    if (rangeIndexes[1] === 0 && rangeIndexes[3] === 0) {
      const sheet: SheetModel = this.spreadsheetObj.getActiveSheet();
      const curRange: string[] = sheet.selectedRange.split(' ');
      const cell: HTMLElement = this.spreadsheetObj.getCell(rangeIndexes[0], 0);
      let cellCheckbox: HTMLInputElement = select('.e-checkbox-wrapper .e-frame', cell);
      if (!cellCheckbox) { return; }
      const checked: boolean = cellCheckbox.classList.contains('e-check');
      args.cancel = true;
      let newRange: string = '';
      curRange.forEach((range: string): void => {
          const curRangeIndexes: number[] = getRangeIndexes(range);
          if (curRangeIndexes[1] === 0 && curRangeIndexes[3] === sheet.colCount - 1) {
            newRange = newRange ? newRange + ' ' + range : range;
          }
      });
      let curSelectRange: string = getCellAddress(rangeIndexes[0], 0) + ':' + getCellAddress(
        rangeIndexes[0], this.spreadsheetObj.getActiveSheet().colCount - 1);
      detach(cellCheckbox.parentElement);
      if (checked) {
        cellCheckbox = createCheckBox(this.spreadsheetObj.createElement, false);
        if (newRange.includes(curSelectRange)) {
          if (newRange.includes(curSelectRange + ' ')) {
            curSelectRange = curSelectRange + ' ';
          } else if (newRange.includes(' ' + curSelectRange)) {
            curSelectRange = ' ' + curSelectRange;
          }
          newRange = newRange.replace(curSelectRange, '');
          if (!newRange) {
            cell.appendChild(cellCheckbox);
            args.cancel = false; return;
          }
        } else {
          cell.appendChild(cellCheckbox);
          args.cancel = false; return;
        }
      } else {
        cellCheckbox = createCheckBox(this.spreadsheetObj.createElement, false, { checked: true });
        newRange = newRange ? newRange + ' ' + curSelectRange : curSelectRange;
      }
      cell.appendChild(cellCheckbox);
      this.spreadsheetObj.selectRange(newRange);
      this.bulkActionMenu.element.disabled = !sheet.selectedRange.includes(' ');
    } else if (rangeIndexes[1] !== 0) {
      requestAnimationFrame((): void => {
        const checkboxes: HTMLElement[] = [].slice.call(selectAll( '.e-checkbox-wrapper .e-frame.e-check', this.spreadsheetObj.element ));
        let cell: Element;
        checkboxes.forEach((checkbox: HTMLElement): void => {
          cell = closest(checkbox, '.e-cell');
          detach(checkbox.parentElement);
          cell.appendChild(createCheckBox(this.spreadsheetObj.createElement, false));
        });
      });
      if (this.autoFillAddress) {
        if (this.curAddress && this.curAddress === args.range) {
          return;
        }
        if (!this.curAddress && (rangeIndexes[0] !== this.autoFillAddress[0] || rangeIndexes[1] !== this.autoFillAddress[1])) {
          args.cancel = true;
          this.curAddress = getRangeAddress(this.autoFillAddress);
          this.spreadsheetObj.selectRange(this.curAddress);
        } else if ( rangeIndexes[2] !== this.autoFillAddress[0] && rangeIndexes[3] !== this.autoFillAddress[1] &&
          rangeIndexes[2] !== this.autoFillAddress[2] && rangeIndexes[3] !== this.autoFillAddress[3]) {
          args.cancel = true;
          const rowDiff: number = Math.abs(this.autoFillAddress[2] - rangeIndexes[2]);
          const colDiff: number = Math.abs(this.autoFillAddress[3] - rangeIndexes[3]);
          let curRange: number[];
          if (rowDiff >= colDiff) {
            curRange = [this.autoFillAddress[0], this.autoFillAddress[1], rangeIndexes[2], this.autoFillAddress[3]];
          } else {
            curRange = [this.autoFillAddress[0], this.autoFillAddress[1], this.autoFillAddress[2], rangeIndexes[3]];
          }
          this.curAddress = getRangeAddress(curRange);
          this.spreadsheetObj.selectRange(this.curAddress);
        } else if (this.curAddress && args.range !== getRangeAddress(this.autoFillAddress)) {
          const swapRange: number[] = getSwapRange(this.autoFillAddress);
          if (rangeIndexes[2] >= swapRange[0] && rangeIndexes[2] <= swapRange[2] && rangeIndexes[3] >= swapRange[1] && rangeIndexes[3] <= swapRange[3]) {
            args.cancel = true;
            this.spreadsheetObj.selectRange(getRangeAddress(this.autoFillAddress));
          } else if (rangeIndexes[3] >= swapRange[1] && rangeIndexes[3] <= swapRange[3] && swapRange[1] !== swapRange[3]) {
            let newRange: string;
            if (rangeIndexes[2] > rangeIndexes[0]) {
              newRange = getRangeAddress([swapRange[0], swapRange[1], rangeIndexes[2], swapRange[3]]);
            } else {
              newRange = getRangeAddress([swapRange[2], swapRange[3], rangeIndexes[2], swapRange[1]]);
            }
            args.cancel = true;
            this.curAddress = newRange;
            this.spreadsheetObj.selectRange(newRange);
          }
        }
      }
    }
  }

  bulkActionSelect(args: MenuEventArgs): void {
    if (args.item.text === 'Bulk Edit') {
      this.bulkEditDlgOpen();
    }
  }
  bulkEditDlgOpen() {
    this.columnData = [];
    const selectIndexes: number[] = getRangeIndexes(this.spreadsheetObj.getActiveSheet().selectedRange);
    const cells: CellModel[] = this.spreadsheetObj.getActiveSheet().rows[0].cells;
    for (let i: number = selectIndexes[1]; i < selectIndexes[3]; i++) {
      if (cells[i] && cells[i].value) {
        this.columnData.push({ Name: cells[i].value });
      }
    }
    this.columnValue = (this.columnData[0] as any).Name;
    this.bulkEditDlgObj.show();
  }
  bulkEditDlgClose() {
    this.bulkEditDlgObj.hide();
  }
  onBultEditColumnChange(args: any): void {
    this.columnValue = args.value;
  }
  bulkEditSave() {
    const index: number =
      this.columnData.findIndex((data: { Name: string }) => data.Name === this.columnValue) + 1;
    if (index > 0) {
      const selectedRanges: string[] = this.spreadsheetObj.getActiveSheet().selectedRange.split(' ');
      const editedValue: string = select('.e-dlg-content .e-textbox .e-input', this.bulkEditDlgObj.element).value;
      let address: string;
      selectedRanges.forEach((selectedRange: string): void => {
          const selectedIndexes: number[] = getRangeIndexes(selectedRange);
          for (let i: number = selectedIndexes[0]; i <= selectedIndexes[2]; i++) {
            address = getCellAddress(i, index);
            this.spreadsheetObj.updateCell({ value: editedValue }, address);
            if (index === 2) {
              this.updateDependantValidation(editedValue, i, index, true);
            }
            this.spreadsheetObj.addInvalidHighlight(address);
          }
        }
      );
      if (index === 1) {
        this.spreadsheetObj.resize();
      }
    }
    this.bulkEditDlgObj.hide();
  }

  onSortComplete(): void {
    this.spreadsheetObj.resize();
  }

  setAutoFill() {
    let values: string[];
    let value: string;
    let patterns: any[];
    const sheet: SheetModel = this.spreadsheetObj.getActiveSheet();
    const selectRange: number[] = getRangeIndexes(sheet.selectedRange);
    let k: number;
    let plen: number;
    let patrn: any;
    let val: number;
    let rowIdx: number;
    const isRFill: boolean = selectRange[0] > selectRange[2];
    const range: number[] = getSwapRange(this.autoFillAddress);
    let clen: number = isRFill ? range[0] - selectRange[2] : selectRange[2] - range[2];
    for (let i: number = range[1]; i <= range[3]; i++) {
      values = [];
      for (let j: number = range[0]; j <= range[2]; j++) {
        value = getCell(j, i, sheet, false, true).value;
        if (value) { values.push(value); }
      }
      patterns = this.getPattern(values, isRFill);
      if (!patterns.length) { continue; }
      k = 0;
      plen = patterns.length;
      while (k < clen) {
        patrn = patterns[k % plen];
        val = null;
        if (isNumber(patrn)) patrn = patterns[patrn];
        switch (patrn.type) {
          case 'number':
            val = this.round(patrn.regVal.a + patrn.regVal.b * patrn.i, 5);
            if (isRFill) {
              patrn.i--;
            } else {
              patrn.i++;
            }
            break;
          case 'string':
            val = patrn.val[patrn.i % patrn.val.length];
            patrn.i++;
            break;
        }
        if (val === null) { continue; }
        if (isRFill) {
          rowIdx = range[0] - 1 - k;
        } else {
          rowIdx = range[2] + 1 + k;
        }
        this.spreadsheetObj.updateCell({ value: val.toString() }, getCellAddress(rowIdx, i));
        k++;
      }
    }
  }

  getPattern(values: string[], isRFill: boolean): any[] {
    let patrns: { type?: string; value?: string[] }[];
    let patrn: { type?: string; value?: string[] };
    let pattern: any[] = [];
    let temp: { regVal?: { a: number; b: number }, val?: string[], type: string, i: number };
    let k: number;
    let l: number;
    let idx: number;
    let len: number;
    let diff: number;
    let regVal: { a: number; b: number };
    patrns = this.getDataPattern(values);
    if (patrns.length) {
      k = 0;
      while (k < patrns.length) {
        patrn = patrns[k];
        switch (patrn.type) {
          case 'number':
            idx = pattern.length;
            len = patrn.value.length;
            diff = isRFill ? -1 : len;
            if (len === 1) {
              patrn.value.push((Number(patrn.value[0]) + 1).toString());
            }
            regVal = this.getPredictionValue(patrn.value);
            temp = { regVal: regVal, type: patrn.type, i: diff };
            pattern.push(temp);
            l = 1;
            while (l < len) {
              pattern.push(idx); l++;
            }
            break;
          case 'string':
            idx = pattern.length;
            temp = { val: patrn.value, type: patrn.type, i: 0 };
            pattern.push(temp);
            l = 1;
            len = patrn.value.length;
            while (l < len) {
              pattern.push(idx); l++;
            }
            break;
        }
        k++;
      }
    }
    return pattern;
  }

  getPredictionValue(args: string[]): { a: number; b: number } {
    let i: number = 0;
    let sumx: number = 0;
    let sumy: number = 0;
    let sumxy: number = 0;
    let sumxx: number = 0;
    let n: number = args.length;
    while (i < n) {
      sumx = sumx + i;
      sumy = sumy + Number(args[i]);
      sumxy = sumxy + i * Number(args[i]);
      sumxx = sumxx + i * i;
      i++;
    }
    const a: number = this.round((sumy * sumxx - sumx * sumxy) / (n * sumxx - sumx * sumx), 5);
    const b: number = this.round((n * sumxy - sumx * sumy) / (n * sumxx - sumx * sumx), 5);
    return { a: a, b: b };
  }

  round(val: number, digits: number) {
    return Number(Math.round(val + 'e' + digits) + 'e-' + digits);
  }

  getDataPattern(data: string[]): { type?: string; value?: string[] }[] {
    let patrn: { type?: string; value?: string[] }[] = [];
    let obj: { type?: string; value?: string[] } = {};
    let val: string;
    let type: string;
    let i: number = 0;
    while (i < data.length) {
      val = data[i];
      type = isNumber(val) ? 'number' : 'string';
      if (i === 0) {
        obj = { value: [val], type: type };
      } else if (type === obj.type) {
        obj.value.push(val);
      } else {
        patrn.push(obj);
        obj = { value: [val], type: type };
      }
      i++;
    }
    patrn.push(obj);
    return patrn;
  }
}
