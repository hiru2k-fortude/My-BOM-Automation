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
    let Bom_Section: string;
    let Season: string;
    let Color_Base: string;
    let Item_Name: string;
    let Style_Number: string;
    //let Width: number;
    // let Length: number;
    //let Size: number;
    //let Diameter: number;
    //let Zipper_Length: number;
    // let WD_Number: number;
    // let Graphic: number;

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
    // let width_check: boolean = false;
    //  let length_check: boolean = false;
    // let size_check: boolean = false;
    // let diameter_check: boolean = false;
    // let zipper_lengthCheck: boolean = false;
    // let wd_numberCheck: boolean = false;
    // let graphic_check: boolean = false;

    // Function to count the number of true conditions
    function countTrueConditions(conditions) {
      // Check if conditions is defined and is an array
      if (!Array.isArray(conditions)) {
        // Handle the case where conditions is not a valid array
        console.error('Invalid conditions array:', conditions);
        return 0; // Return an appropriate value or throw an error
      }
      // Continue with the function logic
      return conditions.filter((condition) => condition).length;
    }
    if (Styles && Array.isArray(Styles)) {
      for (let i = 1; i < Styles.length; i++) {
        let sub_mat_array_under_one_ref = [];
        let Style = Styles[i];
        // go through the rows of excel sheet
        AllData[Style].forEach((ele: any) => {
          //to compare values

          supplierName = ele.SUPPLIER_NAME
            ? ele.SUPPLIER_NAME.toUpperCase().trim()
            : '';
          Supplier_Quality_Reference = ele.SUPPLIER_REF
            ? ele.SUPPLIER_REF.toUpperCase().trim()
            : '';
          Season = ele.SEASON;
          Color_Base = ele.CLR_DYE_TECH;
          Item_Name = ele.ITEM_NAME;

          //get bom section to idntify washes and embelishments
          Bom_Section = ele.BOM_SECTION;

          //Width = ele.
          //Length = ele.
          //Size = ele.
          //Diameter = ele.
          //Zipper_Length = ele.
          //WD_Number = ele.
          //Graphic = ele.

          Style_Number = ele.STYLE_NO_INDIVDUAL;

          // materialType = ele.BOM_SECTION;
          const BrandixQuoteArrays = {};
          const maxConditionCounts = {};

          //washes and finishes , embelishments and grapphics
          const sub_materials_without_supplier_references = [
            'Dye',
            'Wash',
            'Embroidery',
            'Heat Transfer',
            'Print',
            'Applique',
            'Bonding',
          ];

          // Define a function to push quotes into the appropriate array
          function pushToBrandixQuoteArray(
            subMaterialType,
            conditionCount,
            quote,
            ItemName,
            RefNo,
            SupName,
          ) {
            const arrayName = `${subMaterialType}_${conditionCount}`;
            if (!BrandixQuoteArrays[arrayName]) {
              BrandixQuoteArrays[arrayName] = [];
            }
            BrandixQuoteArrays[arrayName].push(quote, ItemName, RefNo, SupName);

            // Update the max condition count for the sub-material type
            if (
              !maxConditionCounts[subMaterialType] ||
              maxConditionCounts[subMaterialType] < conditionCount
            ) {
              maxConditionCounts[subMaterialType] = conditionCount;
            }
          }

          // Define a function to count unique quotes in an array
          function getUniqueQuotes(quotes) {
            return [...new Set(quotes)];
          }

          // Function to process quotes for each sub-material type and condition
          function processQuotes(subMaterialType, conditions, quote) {
            const trueConditionCount = countTrueConditions(conditions);
            BrandixQoute = quote.barndix_quote;
            let ItemName = quote.item_name;
            let RefNo = quote.Supplier_Quality_Reference;
            let SupName = quote.supplier_name;
            uom = quote.uom;
            pushToBrandixQuoteArray(
              subMaterialType,
              trueConditionCount,
              BrandixQoute,
              ItemName,
              RefNo,
              SupName,
            );
          }

          console.log(BrandixQuoteData[Style]);
          console.log('hi hi', BrandixQuoteData[Style].data);

          if (BrandixQuoteData[Style] !== undefined) {
            for (let i = 0; i < BrandixQuoteData[Style].data.length; i++) {
              console.log(BrandixQuoteData[Style].data.length);
              let Qoute = BrandixQuoteData[Style].data[i];

              Supplier_Quality_ReferenceCheck =
                Supplier_Quality_Reference ===
                (Qoute['Supplier_Quality_Reference']
                  ? Qoute['Supplier_Quality_Reference'].toUpperCase()
                  : '');

              if (
                Supplier_Quality_ReferenceCheck == true ||
                Bom_Section == 'Print'
              ) {
                let long_quote_season = StylesPlmData[Style].fs;

                // Extract the season part from e.fs
                const match = long_quote_season.match(/-(\w+)-\d{4}$/);

                if (match) {
                  // Extract and format the season and year
                  const seasonPart = match[0]; // Get the matched part: "-Fall-2024"
                  quote_season = seasonPart
                    .split('-')
                    .filter(Boolean)
                    .join(' '); // Split by hyphen, remove empty strings, and join with a space
                }

                //extreact the sub-material part from item_name of plm data
                subMaterialType = Qoute.item_name;
                let delimiter = ' - ';
                let index = subMaterialType.indexOf(delimiter);
                let extracted_subMaterialType;
                if (index !== -1 && delimiter) {
                  extracted_subMaterialType = subMaterialType.substring(
                    index + 1 + delimiter.length,
                  );
                } else {
                  console.error('index or delimiter is undefined.');
                }
                sub_mat_array_under_one_ref.push(extracted_subMaterialType);

                // Compare with PLM Data
                supplier_nameCheck = Qoute.supplier_name
                  ? Qoute.supplier_name.toUpperCase() === supplierName
                  : false;

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
                  //   size_check,
                ];
                const conditions_Interlining_Pocketing_Warp_Knit_Weft_Knit_Woven =
                  [
                    supplier_nameCheck,
                    Supplier_Quality_ReferenceCheck,
                    company_seasonCheck,
                    color_baseCheck,
                  ];
                const conditions_Button_Eyelet = [
                  supplier_nameCheck,
                  Supplier_Quality_ReferenceCheck,
                  company_seasonCheck,
                  // diameter_check,
                ];
                const conditions_Elastic = [
                  supplier_nameCheck,
                  Supplier_Quality_ReferenceCheck,
                  customer_referenceCheck,
                  // width_check,
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
                  // length_check,
                  // width_check,
                ];
                const conditions_Mobilon = [
                  supplier_nameCheck,
                  Supplier_Quality_ReferenceCheck,
                  company_seasonCheck,
                  // width_check,
                ];
                const conditions_Ring_Slide_Tape = [
                  supplier_nameCheck,
                  Supplier_Quality_ReferenceCheck,
                  company_seasonCheck,
                  // size_check,
                ];
                const conditions_Zipper = [
                  supplier_nameCheck,
                  Supplier_Quality_ReferenceCheck,
                  company_seasonCheck,
                  // zipper_lengthCheck,
                ];
                const conditions_Dye_Wash = [
                  // graphic_check,
                  // wd_numberCheck,
                  company_seasonCheck,
                ];
                const conditions_Embroidery_Heat_Transfer_Print_Applique_Bonding =
                  [company_seasonCheck];
                const conditions_Tag_Pin = [
                  supplier_nameCheck,
                  Supplier_Quality_ReferenceCheck,
                  company_seasonCheck,
                  // length_check,
                ];

                if (
                  extracted_subMaterialType == 'Collar' ||
                  extracted_subMaterialType == 'Cuff'
                ) {
                  processQuotes('Collar_Cuff', conditions_collar_cuff, Qoute);
                }

                if (
                  [
                    'Interlining',
                    'Pocketing',
                    'Warp Knit',
                    'Weft Knit',
                    'Woven',
                  ].includes(extracted_subMaterialType)
                ) {
                  processQuotes(
                    'Interlining_Pocketing_Warp_Knit_Weft_Knit_Woven',
                    conditions_Interlining_Pocketing_Warp_Knit_Weft_Knit_Woven,
                    Qoute,
                  );
                }

                if (
                  extracted_subMaterialType == 'Button' ||
                  extracted_subMaterialType == 'Eyelet'
                ) {
                  processQuotes(
                    'Button_Eyelet',
                    conditions_Button_Eyelet,
                    Qoute,
                  );
                }

                if (extracted_subMaterialType == 'Elastic') {
                  processQuotes('Elastic', conditions_Elastic, Qoute);
                }

                if (
                  [
                    'Heat Seal',
                    'Patch',
                    'Box',
                    'Carton',
                    'Divider',
                    'Hanger',
                    'Poly Bag',
                    'Sticker',
                    'Tag',
                  ].includes(extracted_subMaterialType)
                ) {
                  processQuotes(
                    'Heat_Seal_Patch_Box_Carton_Divider_Hanger_Poly_Bag_Sticker_Tag',
                    conditions_Heat_Seal_Patch_Box_Carton_Divider_Hanger_Poly_Bag_Sticker_Tag,
                    Qoute,
                  );
                }

                if (
                  extracted_subMaterialType == 'Dye' ||
                  extracted_subMaterialType == 'Wash'
                ) {
                  processQuotes('Dye_Wash', conditions_Dye_Wash, Qoute);
                }

                if (
                  [
                    'Embroidery',
                    'Heat Transfer',
                    'Print',
                    'Applique',
                    'Bonding',
                  ].includes(extracted_subMaterialType)
                ) {
                  processQuotes(
                    'Embroidery_Heat_Transfer_Print_Applique_Bonding',
                    conditions_Embroidery_Heat_Transfer_Print_Applique_Bonding,
                    Qoute,
                  );
                }

                if (extracted_subMaterialType == 'Tag Pin') {
                  processQuotes('Tag_Pin', conditions_Tag_Pin, Qoute);
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

          ele.Bom = Bom;

          // Extract and display unique quotes for each sub-material type with the maximum conditions satisfied
          for (const subMaterialType in maxConditionCounts) {
            const maxConditionCount = maxConditionCounts[subMaterialType];
            const arrayName = `${subMaterialType}_${maxConditionCount}`;
            const uniqueQuotes = getUniqueQuotes(BrandixQuoteArrays[arrayName]);
            ele.BrandixQuote = uniqueQuotes;
          }

          // Reset
          BrandixQoute = '';
          uom = '';
        });
      }
    } else {
      console.error('Styles array is undefined or not an array.');
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

    // let BomDetails = await instance.get(
    //   `${Settings.PLM_API}/GetBomItems/${ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]}`,
    // );

    const promises = ExcelData.map(async (ele: any) => {
      if (!AllData[ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]]) {
        AllData[ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]] = [];
        Styles.push(ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]);

        if (
          ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL] != Settings.StyleHeader
        ) {
          // let passingVal = `${Settings.PLM_API}/GetAllStylesAndBoms/${ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]}`;
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
                  // if (s.node_name === Settings.BomDefaultName) {
                  BomPlmData[ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]] = s;

                  BrandixQuoteData[
                    ele[Settings.ExcelIndex.STYLE_NO_INDIVDUAL]
                  ] = (
                    await instance.get(
                      `${Settings.PLM_API}/GetBrandixQuote/${s.latest_revision}`,
                    )
                  ).data;
                  //  }
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
      // BomPlmData: BomPlmData,
      BrandixQuoteData: BrandixQuoteData,
    }));
  }
}
