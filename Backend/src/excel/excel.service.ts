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

    let uom: string = '';
    let quote_season: string;

    let materialType: string = '';

    let Department: string;
    let Division: string;
    let Bom: string;

    let supplier_nameCheck: boolean;
    let Supplier_Quality_ReferenceCheck: boolean;
    let company_seasonCheck: boolean;
    let color_baseCheck: boolean;
    let customer_referenceCheck: boolean;
    let style_noCheck: boolean;

    for (let i = 1; i < Styles.length; i++) {
      let Style = Styles[i];
      // go through the rows of excel sheet
      AllData[Style].forEach((ele: any) => {
        //to compare values
        supplierName = ele.SUPPLIER_NAME.toUpperCase();
        Supplier_Quality_Reference = ele.SUPPLIER_REF.toUpperCase();
        Season = ele.SEASON;
        // Color_Base = ele.CLR_DYE_TECH;
        // Item_Name = ele.ITEM_NAME;
        //no need to compare the style
        Style_Number = ele.STYLE_NO_INDIVDUAL;

        ///////
        //console.log('THIS IS ELE DATA', ele);

        // materialType = ele.BOM_SECTION;

        //if(BrandixQuote[Style] !== undefined){
        let BrandixQuoteArray_1 = [];
        let BrandixQuoteArray_2 = [];
        let BrandixQuoteArray_3 = [];

        for (let i = 0; i < BrandixQuoteData[Style].data.length; i++) {
          let Qoute = BrandixQuoteData[Style].data[i];
          let long_quote_season = StylesPlmData[Style].fs;

          // Extract the season part from e.fs
          const match = long_quote_season.match(
            /-(Spring|Summer|Fall|Winter)-\d{4}$/,
          );

          if (match) {
            // Extract and format the season and year
            const seasonPart = match[0]; // Get the matched part: "-Fall-2024"
            quote_season = seasonPart.split('-').filter(Boolean).join(' '); // Split by hyphen, remove empty strings, and join with a space
          }

          //console.log(quote_season);
          //console.log('this is the style', Style);
          //console.log('THIS IS QUOTE DATA', Qoute);
          materialType = Qoute.material_type;

          if (
            materialType == 'Fabric' ||
            materialType == 'Sewing' ||
            materialType == 'Packing'
          ) {
            // Compare with PLM Data
            supplier_nameCheck =
              Qoute.supplier_name.toUpperCase() === supplierName;

            //console.log('quote supplier',Qoute.supplier_name.toUpperCase());
            //console.log('2nd supplier',supplierName);

            Supplier_Quality_ReferenceCheck =
              Supplier_Quality_Reference ===
              Qoute['Supplier_Quality_Reference'].toUpperCase();

            company_seasonCheck = Season === quote_season;

            if (
              supplier_nameCheck &&
              Supplier_Quality_ReferenceCheck &&
              company_seasonCheck
            ) {
              BrandixQoute = Qoute.barndix_quote;
              BrandixQuoteArray_1.push(BrandixQoute);
              uom = Qoute.uom;
              //console.log('comb 3 quote', BrandixQoute);
            } else if (supplier_nameCheck && Supplier_Quality_ReferenceCheck) {
              BrandixQoute = Qoute.barndix_quote;
              BrandixQuoteArray_2.push(BrandixQoute);
              uom = Qoute.uom;
              // console.log('comb 2-1 quote', BrandixQoute);
            } else if (supplier_nameCheck && company_seasonCheck) {
              BrandixQoute = Qoute.barndix_quote;
              BrandixQuoteArray_2.push(BrandixQoute);
              uom = Qoute.uom;
              // console.log('comb 2-2 quote', BrandixQoute);
            } else if (company_seasonCheck && Supplier_Quality_ReferenceCheck) {
              BrandixQoute = Qoute.barndix_quote;
              BrandixQuoteArray_2.push(BrandixQoute);
              uom = Qoute.uom;
              // console.log('comb 2-3 quote', BrandixQoute);
            } else if (supplier_nameCheck) {
              BrandixQoute = Qoute.barndix_quote;
              BrandixQuoteArray_3.push(BrandixQoute);
              uom = Qoute.uom;
              // console.log('comb 1-1 quote', BrandixQoute);
            } else if (Supplier_Quality_ReferenceCheck) {
              BrandixQoute = Qoute.barndix_quote;
              BrandixQuoteArray_3.push(BrandixQoute);
              uom = Qoute.uom;
              // console.log('comb 1-2 quote', BrandixQoute);
            } else if (company_seasonCheck) {
              BrandixQoute = Qoute.barndix_quote;
              BrandixQuoteArray_3.push(BrandixQoute);
              uom = Qoute.uom;
              //  console.log('comb 1-3 quote', BrandixQoute);
            }
          }

          if (
            materialType == 'Washes&Finishes' ||
            materialType == 'Embellishment'
          ) {
            if (company_seasonCheck) {
              BrandixQoute = Qoute.barndix_quote;
              uom = Qoute.uom;
              BrandixQuoteArray_1.push(BrandixQoute);
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
        } //}

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
        if (UniqBrandixQuoteArray_1.length !== 0) {
          ele.BrandixQuote = UniqBrandixQuoteArray_1;
          console.log('this is array 1', UniqBrandixQuoteArray_1);
        } else {
          if (UniqBrandixQuoteArray_2.length !== 0) {
            ele.BrandixQuote = UniqBrandixQuoteArray_2;
            console.log('this is array 2', UniqBrandixQuoteArray_2);
          } else {
            ele.BrandixQuote = UniqBrandixQuoteArray_3;
            console.log('this is array 3', UniqBrandixQuoteArray_3);
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
            const match = e.fs.match(/-(Spring|Summer|Fall|Winter)-\d{4}$/);

            if (match) {
              // Extract and format the season and year
              const seasonPart = match[0]; // Get the matched part: "-Fall-2024"
              const seasonYear = seasonPart
                .split('-')
                .filter(Boolean)
                .join(' '); // Split by hyphen, remove empty strings, and join with a space

              if (seasonYear === ele[Settings.ExcelIndex.SEASON]) {
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
              const match = key.match(/-\s*(Spring|Summer|Fall|Winter)-\d{4}$/);
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
