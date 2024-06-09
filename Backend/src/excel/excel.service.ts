import { Injectable } from '@nestjs/common';
import axios from 'axios';
import { Console } from 'console';

const https = require('https');
const fs = require('fs');
const xlsx = require('xlsx');
const Settings = require('../../setting');

// Create an HTTPS agent that ignores SSL certificate errors
const httpsAgent = new https.Agent({
  rejectUnauthorized: false,
});

// Axios instance with custom HTTPS agent
const instance = axios.create({
  httpsAgent: httpsAgent,
});

@Injectable()
export class ExcelService {
  /**
   * Sets the UEE template data based on the provided form data and UEE data.
   * @param formData - The form data.
   * @param UEEData - The UEE data.
   * @returns The updated UEE template data.
   */
  setUEETemplateData(formData: any, UEEData: any): any {
    const AllData = UEEData.All;
    const Styles = UEEData.Styles;
    const StylesPlmData = UEEData.StylesPlmData;
    const BrandixQuoteData = UEEData.BrandixQuoteData;

    let BrandixQoute: string = '';

    let supplierName: string;
    let Supplier_Quality_Reference: string;
    let Season: string;
    let Color_Base: string;
    let Item_Name: string;
    let Style_Number: string;
    let Width: number;
    let Length: number;
    let Size: number;
    let Diameter: number;
    let Zipper_Length: number;
    let WD_Number: number;
    let Graphic: number;

    let uom: string = '';
    let quote_season: string;

    let subMaterialType: string = '';

    let Department: string;
    let Division: string;
    let Bom: string;

    let supplier_nameCheck: boolean = false;
    let Supplier_Quality_ReferenceCheck: boolean = false;
    let company_seasonCheck: boolean = false;
    let color_baseCheck: boolean = false;
    let customer_referenceCheck: boolean = false; //item name
    let style_noCheck: boolean = false;
    let width_check: boolean = false;
    let length_check: boolean = false;
    let size_check: boolean = false;
    let diameter_check: boolean = false;
    let zipper_lengthCheck: boolean = false;
    let wd_numberCheck: boolean = false;
    let graphic_check: boolean = false;

    // Function to count the number of true conditions
    function countTrueConditions(conditions) {
      return conditions.filter((condition) => condition).length;
    }

    for (let i = 1; i < Styles.length; i++) {
      let Style = Styles[i];
      // go through the rows of excel sheet
      AllData[Style].forEach((ele: any) => {
        //to compare values
        supplierName = ele.SUPPLIER_NAME.toUpperCase();
        Supplier_Quality_Reference = ele.SUPPLIER_REF.toUpperCase();
        Season = ele.SEASON;
        Color_Base = ele.CLR_DYE_TECH;
        Item_Name = ele.ITEM_NAME;
        //Width = ele.
        //Length = ele.
        //Size = ele.
        //Diameter = ele.
        //Zipper_Length = ele.
        //WD_Number = ele.
        //Graphic = ele.

        //
        //no need to compare the style
        Style_Number = ele.STYLE_NO_INDIVDUAL;

        ///////
        //console.log('THIS IS ELE DATA', ele);

        // materialType = ele.BOM_SECTION;
        let BrandixQuoteArray_1 = [];
        let BrandixQuoteArray_2 = [];
        let BrandixQuoteArray_3 = [];
        let BrandixQuoteArray_4 = [];
        let BrandixQuoteArray_5 = [];

        if (BrandixQuoteData[Style] !== undefined) {
          for (let i = 0; i < BrandixQuoteData[Style].data.length; i++) {
            let Qoute = BrandixQuoteData[Style].data[i];
            let long_quote_season = StylesPlmData[Style].fs;

            // Extract the season part from e.fs
            const match = long_quote_season.match(/-(\w+)-\d{4}$/);

            if (match) {
              // Extract and format the season and year
              const seasonPart = match[0]; // Get the matched part: "-Fall-2024"
              quote_season = seasonPart.split('-').filter(Boolean).join(' '); // Split by hyphen, remove empty strings, and join with a space
            }

            //console.log(quote_season);
            //console.log('this is the style', Style);
            //console.log('THIS IS QUOTE DATA', Qoute);

            //extreact the sub-material part from item_name of plm data
            subMaterialType = Qoute.item_name;
            let delimiter = ' - ';
            let index = subMaterialType.indexOf(delimiter);
            let extracted_subMaterialType = subMaterialType.substring(
              index + 1 + delimiter.length,
            );

            // Compare with PLM Data
            supplier_nameCheck =
              Qoute.supplier_name.toUpperCase() === supplierName;

            //console.log('quote supplier',Qoute.supplier_name.toUpperCase());
            //console.log('2nd supplier',supplierName);

            Supplier_Quality_ReferenceCheck =
              Supplier_Quality_Reference ===
              Qoute['Supplier_Quality_Reference'].toUpperCase();

            company_seasonCheck = Season === quote_season;
            //color_baseCheck = Color_Base ===
            //size_check = Size ===
            //customer_referenceCheck ===
            //width_check ===
            // length_check ===
            //diameter_check===
            //wd_numberCheck===
            //graphic_check===
            //zipper_lengthCheck ===

            // List of conditions
            const conditions_collar_cuff = [
              supplier_nameCheck,
              Supplier_Quality_ReferenceCheck,
              company_seasonCheck,
              color_baseCheck,
              size_check,
            ];
            const conditions_Interlining_Pocketing_Warp_Knit_Weft_Knit_Woven = [
              supplier_nameCheck,
              Supplier_Quality_ReferenceCheck,
              company_seasonCheck,
              color_baseCheck,
            ];
            const conditions_Button_Eyelet = [
              supplier_nameCheck,
              Supplier_Quality_ReferenceCheck,
              company_seasonCheck,
              diameter_check,
            ];
            const conditions_Elastic = [
              supplier_nameCheck,
              Supplier_Quality_ReferenceCheck,
              customer_referenceCheck,
              width_check,
            ];
            const conditions_Heat_Seal_Patch_Box_Carton_Divider_Hanger_Poly_Bag_Sticker_Tag =
              [
                supplier_nameCheck,
                Supplier_Quality_ReferenceCheck,
                company_seasonCheck,
              ];
            const conditions_Label = [
              supplier_nameCheck,
              Supplier_Quality_ReferenceCheck,
              company_seasonCheck,
              length_check,
              width_check,
            ];
            const conditions_Mobilon = [
              supplier_nameCheck,
              Supplier_Quality_ReferenceCheck,
              company_seasonCheck,
              width_check,
            ];
            const conditions_Ring_Slide_Tape = [
              supplier_nameCheck,
              Supplier_Quality_ReferenceCheck,
              company_seasonCheck,
              size_check,
            ];
            const conditions_Zipper = [
              supplier_nameCheck,
              Supplier_Quality_ReferenceCheck,
              company_seasonCheck,
              zipper_lengthCheck,
            ];
            const conditions_Dye_Wash = [
              graphic_check,
              wd_numberCheck,
              company_seasonCheck,
            ];
            const conditions_Embroidery_Heat_Transfer_Print_Applique_Bonding = [
              company_seasonCheck,
            ];
            const conditions_Tag_Pin = [
              supplier_nameCheck,
              Supplier_Quality_ReferenceCheck,
              company_seasonCheck,
              length_check,
            ];

            if (
              extracted_subMaterialType == 'Collar' ||
              extracted_subMaterialType == 'Cuff'
            ) {
              // Count the number of true conditions
              const trueConditionCount = countTrueConditions(
                conditions_collar_cuff,
              );
              // Check if at least 4 conditions are true
              if (trueConditionCount == 5) {
                BrandixQoute = Qoute.barndix_quote;
                BrandixQuoteArray_1.push(BrandixQoute);
                uom = Qoute.uom;
                //console.log('comb 3 quote', BrandixQoute);
              } else if (trueConditionCount == 4) {
                BrandixQoute = Qoute.barndix_quote;
                BrandixQuoteArray_2.push(BrandixQoute);
                uom = Qoute.uom;
                // console.log('comb 2-1 quote', BrandixQoute);
              } else if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                BrandixQuoteArray_3.push(BrandixQoute);
                uom = Qoute.uom;
                // console.log('comb 2-2 quote', BrandixQoute);
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                BrandixQuoteArray_4.push(BrandixQoute);
                uom = Qoute.uom;
                // console.log('comb 2-3 quote', BrandixQoute);
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                BrandixQuoteArray_5.push(BrandixQoute);
                uom = Qoute.uom;
                // console.log('comb 1-1 quote', BrandixQoute);
              }
            }

            if (
              extracted_subMaterialType == 'Interlining' ||
              extracted_subMaterialType == 'Pocketing' ||
              extracted_subMaterialType == 'Warp Knit' ||
              extracted_subMaterialType == 'Weft Knit' ||
              extracted_subMaterialType == 'Woven'
            ) {
              // Count the number of true conditions
              const trueConditionCount = countTrueConditions(
                conditions_Interlining_Pocketing_Warp_Knit_Weft_Knit_Woven,
              );
              if (trueConditionCount == 4) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_4.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (
              extracted_subMaterialType == 'Button' ||
              extracted_subMaterialType == 'Eyelet'
            ) {
              // Count the number of true conditions
              const trueConditionCount = countTrueConditions(
                conditions_Button_Eyelet,
              );
              if (trueConditionCount == 4) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_4.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (extracted_subMaterialType == 'Elastic') {
              // Count the number of true conditions
              const trueConditionCount =
                countTrueConditions(conditions_Elastic);
              if (trueConditionCount == 4) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_4.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (
              extracted_subMaterialType == 'Heat Seal' ||
              extracted_subMaterialType == 'Patch' ||
              extracted_subMaterialType == 'Box' ||
              extracted_subMaterialType == 'Carton' ||
              extracted_subMaterialType == 'Divider' ||
              extracted_subMaterialType == 'Hanger' ||
              extracted_subMaterialType == 'Poly Bag' ||
              extracted_subMaterialType == 'Sticker' ||
              extracted_subMaterialType == 'Tag'
            ) {
              // Count the number of true conditions
              const trueConditionCount = countTrueConditions(
                conditions_Heat_Seal_Patch_Box_Carton_Divider_Hanger_Poly_Bag_Sticker_Tag,
              );
              if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (extracted_subMaterialType == 'Label') {
              // Count the number of true conditions
              const trueConditionCount = countTrueConditions(conditions_Label);
              if (trueConditionCount == 5) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 4) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_4.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_5.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (extracted_subMaterialType == 'Mobilon') {
              // Count the number of true conditions
              const trueConditionCount =
                countTrueConditions(conditions_Mobilon);
              if (trueConditionCount == 4) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_4.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (
              extracted_subMaterialType == 'Ring' ||
              extracted_subMaterialType == 'Slide' ||
              extracted_subMaterialType == 'Tape'
            ) {
              // Count the number of true conditions
              const trueConditionCount = countTrueConditions(
                conditions_Ring_Slide_Tape,
              );
              if (trueConditionCount == 4) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_4.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (extracted_subMaterialType == 'Zipper') {
              // Count the number of true conditions
              const trueConditionCount = countTrueConditions(conditions_Zipper);
              if (trueConditionCount == 4) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_4.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (
              extracted_subMaterialType == 'Dye' ||
              extracted_subMaterialType == 'Wash'
            ) {
              // Count the number of true conditions
              const trueConditionCount =
                countTrueConditions(conditions_Dye_Wash);
              if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (
              extracted_subMaterialType == 'Embroidery' ||
              extracted_subMaterialType == 'Heat Transfer' ||
              extracted_subMaterialType == 'Print' ||
              extracted_subMaterialType == 'Applique' ||
              extracted_subMaterialType == 'Bonding'
            ) {
              // Count the number of true conditions
              const trueConditionCount = countTrueConditions(
                conditions_Embroidery_Heat_Transfer_Print_Applique_Bonding,
              );
              if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
            if (extracted_subMaterialType == 'Tag Pin') {
              // Count the number of true conditions
              const trueConditionCount =
                countTrueConditions(conditions_Tag_Pin);
              if (trueConditionCount == 4) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_1.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 3) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_2.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              } else if (trueConditionCount == 2) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the branfdix quote data loop imedietly
              } else if (trueConditionCount == 1) {
                BrandixQoute = Qoute.barndix_quote;
                uom = Qoute.uom;
                BrandixQuoteArray_3.push(BrandixQoute);
                // break; //exit the brandix quote data loop imedietly
              }
            }
          }

          if (StylesPlmData[Style] !== undefined) {
            Department = StylesPlmData[Style].department;
            Division = StylesPlmData[Style].division;
            Bom = Settings.BomDefaultName;

            ele.uom = uom;
            ele.Department = Department;
            ele.Division = Division;
          }
        }

        /* if (BrandixQoute == '') {
          if (
            supplier_nameCheck === false &&
            Supplier_Quality_ReferenceCheck === false
          )
            BrandixQoute = 'Supplier Error, Reference Error';
          else if (supplier_nameCheck === false)
            BrandixQoute = 'Supplier Error';
          else BrandixQoute = 'Reference Error';
        }
        */
        ele.Bom = Bom;

        const UniqBrandixQuoteArray_1 = [];
        const UniqBrandixQuoteArray_2 = [];
        const UniqBrandixQuoteArray_3 = [];
        const UniqBrandixQuoteArray_4 = [];
        const UniqBrandixQuoteArray_5 = [];
        const seen = {};

        BrandixQuoteArray_1.forEach((value) => {
          if (!seen[value]) {
            UniqBrandixQuoteArray_1.push(value);
            seen[value] = true;
          }
        });

        BrandixQuoteArray_2.forEach((value) => {
          if (!seen[value]) {
            UniqBrandixQuoteArray_2.push(value);
            seen[value] = true;
          }
        });

        BrandixQuoteArray_3.forEach((value) => {
          if (!seen[value]) {
            UniqBrandixQuoteArray_3.push(value);
            seen[value] = true;
          }
        });
        BrandixQuoteArray_4.forEach((value) => {
          if (!seen[value]) {
            UniqBrandixQuoteArray_4.push(value);
            seen[value] = true;
          }
        });
        BrandixQuoteArray_5.forEach((value) => {
          if (!seen[value]) {
            UniqBrandixQuoteArray_5.push(value);
            seen[value] = true;
          }
        });
        if (UniqBrandixQuoteArray_1.length !== 0) {
          ele.BrandixQuote = UniqBrandixQuoteArray_1;
          // console.log('this is array 1', UniqBrandixQuoteArray_1);
        } else {
          if (UniqBrandixQuoteArray_2.length !== 0) {
            ele.BrandixQuote = UniqBrandixQuoteArray_2;
            //   console.log('this is array 2', UniqBrandixQuoteArray_2);
          } else {
            if (UniqBrandixQuoteArray_3.length !== 0) {
              ele.BrandixQuote = UniqBrandixQuoteArray_3;
              //  console.log('this is array 3', UniqBrandixQuoteArray_3);
            } else {
              if (UniqBrandixQuoteArray_4.length !== 0) {
                ele.BrandixQuote = UniqBrandixQuoteArray_4;
                //   console.log('this is array 4', UniqBrandixQuoteArray_4);
              } else {
                ele.BrandixQuote = UniqBrandixQuoteArray_5;
                //   console.log('this is array 5', UniqBrandixQuoteArray_5);
              }
            }
          }
        }

        // Reset
        BrandixQoute = '';
        uom = '';
      });
    }

    return AllData;
  }

  /**
   * Reads the data from an Excel sheet.
   * @param file - The Excel file.
   * @param sheetIndex - The index of the sheet to read.
   * @returns The data read from the Excel sheet.
   */
  readExcelSheet(file: any, sheetIndex: number): any {
    try {
      let buffer;
      if (Buffer.isBuffer(file.buffer)) {
        buffer = file.buffer;
      } else {
        try {
          buffer = fs.readFileSync(file);
        } catch (err) {
          console.error('Error reading file:', err);
          return null;
        }
      }

      const workbook = xlsx.read(buffer, { type: 'buffer' });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
      return data;
    } catch (error) {
      console.log(error.message);
      throw error;
    }
  }

  /**
   * Sets the UEE row data based on the provided form data and Excel data.
   * @param formData - The form data.
   * @param ExcelData - The Excel data.
   * @returns A promise that resolves to the updated UEE row data.
   */
  async setUEERowData(formData: any, ExcelData: any): Promise<any> {
    const AllData = {};
    const StylesPlmData = {};
    const BrandixQuoteData = {};

    const BomPlmData = {};
    const Styles = [];

    const promises = ExcelData.map(async (ele: any) => {
      if (!AllData[ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]]) {
        AllData[ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]] = [];
        Styles.push(ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]);

        if (
          ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL] != Settings.StyleHeader
        ) {
          let StyleDetails = await instance.get(
            `${Settings.PLM_API}/GetAllStylesAndBoms/${ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]}`,
          );

          //iterate through season list
          StyleDetails.data.SeasonList.forEach((e: any) => {
            //if fs == season in excel that season data will store in stylesPlmData
            // Extract the season part from e.fs
            const match = e.fs.match(/-(\w+)-\d{4}$/);

            if (match) {
              // Extract and format the season and year
              const seasonPart = match[0]; // Get the matched part: "-Fall-2024"
              const seasonYear1 = match[0].slice(1); // Remove the leading hyphen
              // Replace the hyphen with a space
              const seasonYear2 = seasonYear1.replace('-', ' '); // Convert "Fall-2024" to "Fall 2024"

              // If the formatted season and year matches the expected value
              if (seasonYear2 === ele[Settings.ExcelIndex.SEASON]) {
                StylesPlmData[ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]] = e;
              }
            }
          });
          ///////
          console.log(
            'this is the style details season list',
            StyleDetails.data.SeasonList,
          );
          console.log(
            'this is the style details bom list',
            StyleDetails.data.BomList,
          );

          const bomPromises = StyleDetails.data.BomList.map(async (e: any) => {
            // Extract the season part from e.fs in the BOM list
            const bomSeasonKey = Object.keys(e).find((key) => {
              const match = key.match(/-(\w+)-\d{4}$/);
              if (match) {
                const seasonPart = match[0]; // Get the matched part: "-Fall-2024"
                const formattedSeason = seasonPart
                  .split('-')
                  .filter(Boolean)
                  .join(' '); // Split by hyphen, remove empty strings, and join with a space
                return formattedSeason === ele[Settings.ExcelIndex.SEASON];
              }
              return false;
            });

            if (bomSeasonKey) {
              // If a matching season key was found, process the BOM data for that season
              return Promise.all(
                e[bomSeasonKey].map(async (s: any) => {
                  if (s.node_name === Settings.BomDefaultName) {
                    BomPlmData[ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]] = s;

                    BrandixQuoteData[
                      ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]
                    ] = (
                      await instance.get(
                        `${Settings.PLM_API}/GetBrandixQuote/${s.latest_revision}`,
                      )
                    ).data;
                  }
                }),
              );
            }
          });

          await Promise.all(bomPromises);
        }
      }

      // Append the data
      AllData[ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]].push({
        IDX: ele[Settings.ExcelIndex.IDX],
        SEASON: ele[Settings.ExcelIndex.SEASON],
        CATEGORY: ele[Settings.ExcelIndex.CATEGORY],
        PROGRAM: ele[Settings.ExcelIndex.PROGRAM],
        STYLE_NO_INDIVDUAL: ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL],
        GMT_DESCRIPTION: ele[Settings.ExcelIndex.GMT_DESCRIPTION],
        GMT_COLOR: ele[Settings.ExcelIndex.GMT_COLOR],
        NRF: ele[Settings.ExcelIndex.NRF],
        GMT_COLOR_CODE: ele[Settings.ExcelIndex.GMT_COLOR_CODE],
        PACK_COMBINATION: ele[Settings.ExcelIndex.PACK_COMBINATION],
        PLACEMENT_NAME: ele[Settings.ExcelIndex.PLACEMENT_NAME],
        BOM_SECTION: ele[Settings.ExcelIndex.BOM_SECTION],
        ITEM_NAME: ele[Settings.ExcelIndex.ITEM_NAME],
        SUPPLIER_NAME: ele[Settings.ExcelIndex.SUPPLIER_NAME],
        RM_COLOR_NAME: ele[Settings.ExcelIndex.RM_COLOR_NAME],
        CLR_DYE_TECH: ele[Settings.ExcelIndex.CLR_DYE_TECH],
        RM_COLOR_REF: ele[Settings.ExcelIndex.RM_COLOR_REF],
        GARMENT_WAY: ele[Settings.ExcelIndex.GARMENT_WAY],
        SUPPLIER_REF: ele[Settings.ExcelIndex.SUPPLIER_REF],
        MATERIAL_TYPE: ele[Settings.ExcelIndex.MATERIAL_TYPE],
      });
    });

    return Promise.all(promises).then(() => ({
      All: AllData,
      Styles: Styles,
      StylesPlmData: StylesPlmData,
      BomPlmData: BomPlmData,
      BrandixQuoteData: BrandixQuoteData,
    }));
  }
}
