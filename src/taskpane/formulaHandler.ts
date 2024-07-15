/* global console, Excel, OfficeExtension */

const changingCells = {} as Record<string, boolean>;

class FormulaHandler {
  private events = {} as Record<string, OfficeExtension.EventHandlerResult<any>>;
  /**
   *
   */
  constructor() {}

  public async initializeEvents(): Promise<void> {
    if (typeof Excel == "undefined") return;

    console.log("[EXCEL EVENTS] > INITIALIZING");

    await this.bindEvents();
  }

  public async bindExcelEvents(): Promise<void> {
    if (typeof Excel == "undefined") return;

    await this.bindEvents();
  }

  public async unbindExcelEvents(): Promise<void> {
    if (typeof Excel == "undefined") return;

    await this.unbindEvents();
  }

  private async bindEvents() {
    await Excel.run(async (context) => {
      console.log(`[EXCEL EVENTS] > BINDING EVENTS TO CONTEXT`, context);

      if (this.events["FormulaChanged"] != null) {
        console.error(`[EXCEL EVENTS] > FORMULA CHANGED EVENT ALREADY BOUND`, this.events["FormulaChanged"]);
      } else {
        console.log(`[EXCEL EVENTS] > BINDING FORMULA CHANGED EVENT`);
        this.events["FormulaChanged"] = context.workbook.worksheets.onFormulaChanged.add(this.formulaChangeHandler);
      }

      if (this.events["SelectionChanged"] != null) {
        console.error(`[EXCEL EVENTS] > SELECTION CHANGED EVENT ALREADY BOUND`, this.events["SelectionChanged"]);
      } else {
        console.debug(`[EXCEL EVENTS] > BINDING SELECTION CHANGED EVENT`);
        this.events["SelectionChanged"] = context.workbook.worksheets.onSelectionChanged.add(this.selectionChanged);
      }

      await context.sync();
    });
  }

  private async unbindEvents() {
    for (const eventName in this.events) {
      let event = this.events[eventName];

      // reuse the context from the event as per: https://learn.microsoft.com/en-au/office/dev/add-ins/excel/excel-add-ins-events#remove-an-event-handler
      await Excel.run(event.context, async (context) => {
        try {
          console.log(`[EXCEL EVENTS] > UNBINDING EVENT FROM CONTEXT > ${eventName}`, event, context);

          event.remove();

          await context.sync();

          console.log(`[EXCEL EVENTS] > UNBINDING SUCCESSFUL > ${eventName}`);
        } catch (e) {
          console.error(`[EXCEL EVENTS] > UNBINDING FAILED > BYPASS > ${eventName}`, e);
        } finally {
          // No matter what, clean it up from the event list so it can be re-bound
          delete this.events[eventName];
        }
      });
    }
  }

  private async formulaChangeHandler(e: Excel.WorksheetFormulaChangedEventArgs): Promise<void> {
    await Excel.run(async (context) => {
      const cellAddress = e.formulaDetails[0].cellAddress;

      if (changingCells[cellAddress]) {
        console.error(`[EXCEL EVENTS - FORMULA CHANGED] > ${cellAddress} > Duplicate detected`, e);
        return;
      }

      changingCells[cellAddress] = true;

      try {
        console.log(`[EXCEL EVENTS - FORMULA CHANGED] > ${cellAddress}`, e, context);
      } finally {
        changingCells[cellAddress] = false;
      }
    });
  }

  private async selectionChanged(e: Excel.WorksheetSelectionChangedEventArgs): Promise<void> {
    const cellAddress = e.address;

    if (changingCells[cellAddress]) {
      console.error(`[EXCEL EVENTS - SELECTION CHANGED] > ${cellAddress} > Duplicate detected`, e);
      return;
    }

    changingCells[cellAddress] = true;

    try {
      console.debug(`[EXCEL EVENTS - SELECTION CHANGED] > ${cellAddress} `, e, e.address, e.worksheetId, e.type);
    } finally {
      changingCells[cellAddress] = false;
    }
  }
}

export const formulaHandler = new FormulaHandler();
