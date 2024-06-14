import {
  Controller,
  Post,
  UseInterceptors,
  Res,
  UploadedFiles,
  Body,
} from '@nestjs/common';
import { ExcelService } from './excel.service';
import { AnyFilesInterceptor } from '@nestjs/platform-express';
import { Response } from 'express';
import { Console } from 'console';

/**
 * Uploads a file and processes it to generate UEE template data.
 * @param files - The uploaded files.
 * @param formData - The form data.
 * @param res - The response object.
 */
@Controller('excel')
export class ExcelController {
  constructor(private readonly excelService: ExcelService) {}

  @Post()
  @UseInterceptors(AnyFilesInterceptor())
  async uploadFile(
    @UploadedFiles() files: Array<Express.Multer.File>,
    @Body() formData: any,
    @Res() res: Response,
  ) {
    let TemplateData: any;
    try {
      // Read the Excel sheet
      const data = this.excelService.readExcelSheet(files[0], 0);

      // Set UEE row data
      const UEEData = await this.excelService.setUEERowData(formData, data);

      // Set UEE template data
      TemplateData = await this.excelService.setUEETemplateData(
        formData,
        UEEData,
      );

      // Send the UEE template data as response
      res.send({ UEETemplate: TemplateData });
    } catch (error) {
      console.log('\n  Error in 1st time :: ' + error.message);
      console.log(
        '\t Pleace wait...' +
          ' -- We are currently reprocessing the request --',
      );
      try {
        // Retry the process
        const data = this.excelService.readExcelSheet(files[0], 0);
        const UEEData = await this.excelService.setUEERowData(formData, data);
        TemplateData = await this.excelService.setUEETemplateData(
          formData,
          UEEData,
        );
        console.log('Success 2nd Time');
        res.send({ UEETemplate: TemplateData });
      } catch (error) {
        console.log('\n  Error in 2nd time :: ' + error.message);
        console.log(
          '\t Pleace wait...' +
            ' -- We are currently reprocessing the request --',
        );
        try {
          // Retry the process again
          const data = this.excelService.readExcelSheet(files[0], 0);
          const UEEData = await this.excelService.setUEERowData(formData, data);
          TemplateData = await this.excelService.setUEETemplateData(
            formData,
            UEEData,
          );
          console.log('Success 3rd Time');
          res.send({ UEETemplate: TemplateData });
        } catch (error) {
          console.log('\n  Error in 3rd time :: ' + error.message);
          console.log(
            '\t Pleace wait...' +
              ' -- We are currently reprocessing the request --',
          );
          try {
            // Retry the process again
            const data = this.excelService.readExcelSheet(files[0], 0);
            const UEEData = await this.excelService.setUEERowData(
              formData,
              data,
            );
            TemplateData = await this.excelService.setUEETemplateData(
              formData,
              UEEData,
            );
            console.log('Success 4th Time');
            res.send({ UEETemplate: TemplateData });
          } catch (error) {
            // If all retries fail, send an error response
            res.status(500).send({ message: error.message });
          }
        }
      }
    }
  }
}
