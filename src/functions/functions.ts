/** @CustomFunction
 * @description Increments the cell with a given amount at a specified interval in milliseconds.
 * @param {any[][]} arg1 - The amount to add to the cell value on each increment.
 * @param {any[][]} arg2 - The time in milliseconds to wait before the next increment on the cell.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation - Parameter to send results to Excel
 *     or respond to the user canceling the function.
 * @returns An incrementing value.
 */
export function fnMain(arg1: any[][], arg2: any[][], invocation: CustomFunctions.StreamingInvocation<any[][]>): void {
  invocation.setResult([["Processing"]]);
}

/** @CustomFunction
 * @description Increments the cell with a given amount at a specified interval in milliseconds.
 * @param {any[][]} arg1 - The amount to add to the cell value on each increment.
 * @param {any[][]} arg2 - The time in milliseconds to wait before the next increment on the cell.
 * @param {CustomFunctions.StreamingInvocation<number>} invocation - Parameter to send results to Excel
 *     or respond to the user canceling the function.
 * @returns any[][] An incrementing value.
 */
export function fnEntity(arg1: any[][], invocation: CustomFunctions.StreamingInvocation<any[][]>): void {
  const value1 = arg1[0][0].toString();

  const result = {
    type: "Entity",
    basicType: "Error",
    basicValue: "#VALUE!",
    text: value1,
    properties: {
      CashflowCategory: {
        type: "String",
        basicType: "String",
        basicValue: "PL",
      },
      childrenCardinality: {
        type: "String",
        basicType: "String",
        basicValue: "0",
        propertyMetadata: {
          excludeFrom: {
            autoComplete: true,
            calcCompare: true,
            cardView: true,
            dotNotation: true,
          },
        },
      },
      Class: {
        type: "String",
        basicType: "String",
        basicValue: "",
        propertyMetadata: {
          excludeFrom: {
            cardView: true,
          },
        },
      },
      code: {
        type: "String",
        basicType: "String",
        basicValue: "200",
        propertyMetadata: {
          excludeFrom: {
            calcCompare: true,
            cardView: true,
          },
        },
      },
      Description: {
        type: "String",
        basicType: "String",
        basicValue: "",
        propertyMetadata: {
          excludeFrom: {
            cardView: true,
          },
        },
      },
      hierarchy: {
        type: "String",
        basicType: "String",
        basicValue: "Account",
        propertyMetadata: {
          excludeFrom: {
            calcCompare: true,
            cardView: true,
          },
        },
      },
      hierarchyKey: {
        type: "String",
        basicType: "String",
        basicValue: "8",
        propertyMetadata: {
          excludeFrom: {
            autoComplete: true,
            calcCompare: true,
            cardView: true,
            dotNotation: true,
          },
        },
      },
      key: {
        type: "String",
        basicType: "String",
        basicValue: "101775",
        propertyMetadata: {
          excludeFrom: {
            cardView: true,
          },
        },
      },
      Mappings: {
        type: "String",
        basicType: "String",
        basicValue: "BobCo",
      },
      memberType: {
        type: "String",
        basicType: "String",
        basicValue: "L",
        propertyMetadata: {
          excludeFrom: {
            autoComplete: true,
            calcCompare: true,
            cardView: true,
            dotNotation: true,
          },
        },
      },
      name: {
        type: "String",
        basicType: "String",
        basicValue: value1,
        propertyMetadata: {
          excludeFrom: {
            calcCompare: true,
            cardView: true,
          },
        },
      },
      "Reporting Scale": {
        type: "FormattedNumber",
        basicType: "Double",
        basicValue: -1,
        numberFormat: "#",
      },
      sortOrder: {
        type: "String",
        basicType: "String",
        basicValue: "8",
        propertyMetadata: {
          excludeFrom: {
            autoComplete: true,
            calcCompare: true,
            cardView: true,
            dotNotation: true,
          },
        },
      },
      "Tax Type": {
        type: "String",
        basicType: "String",
        basicValue: "NONE",
      },
      TaxType: {
        type: "String",
        basicType: "String",
        basicValue: "",
        propertyMetadata: {
          excludeFrom: {
            cardView: true,
          },
        },
      },
      XeroAccountCode: {
        type: "String",
        basicType: "String",
        basicValue: "200",
      },
    },
    layouts: {
      card: {
        layout: "Entity",
        title: value1,
        subTitle: {
          property: "code",
        },
        sections: [
          {
            layout: "List",
            title: "Additional information",
            collapsed: false,
            properties: [
              "Description",
              "Description",
              "Tax Type",
              "Mappings",
              "Reporting Scale",
              "CashflowCategory",
              "Class",
              "TaxType",
              "XeroAccountCode",
            ],
          },
        ],
      },
      compact: {
        icon: "Layer",
      },
    },
    provider: {
      description: "powered by xpna",
      logoSourceAddress: "https://xpna.app/assets/xpna-logo.svg",
      logoTargetAddress: "https://xpna.co",
    },
  };

  invocation.setResult([[result]]);
}
