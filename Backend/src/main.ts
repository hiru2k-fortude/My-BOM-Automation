// main.ts
import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';
import { NestExpressApplication } from '@nestjs/platform-express';
const path = require('path');

const Setting = require('../setting');

async function bootstrap() {
  const app = await NestFactory.create<NestExpressApplication>(AppModule);
  app.enableCors(); // Enable CORS if needed
  await app.listen(Setting.port);
  console.log(`Server is running on: http://localhost:${Setting.port}`);
  // console.log(Setting);
  
// Using __dirname
// const projectRootDir1 = __dirname;

// // Using process.cwd()
// const projectRootDir2 = process.cwd();

// // Resolving to an absolute path
// const projectRootDirAbsolute = path.resolve(__dirname);

// console.log("Using __dirname:", projectRootDir1);
// console.log("Using process.cwd():", projectRootDir2);
// console.log("Resolved absolute path:", projectRootDirAbsolute);
}
bootstrap();
