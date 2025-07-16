import * as http from 'http';

import { NestFactory } from '@nestjs/core';
import { AppModule } from './app.module';

async function bootstrap() {
  const app = await NestFactory.create(AppModule);

  // Enable CORS
  app.enableCors({
    origin: 'https://internal-invocie.opendreamcorp.com', // Your Vite frontend URL
    methods: ['*'],
    allowedHeaders: ['*'],
    credentials: true,
  });

  app.setGlobalPrefix('api');

  const server = app.getHttpServer() as http.Server;
  server.setTimeout(10 * 60 * 1000);

  await app.listen(8000);
}
bootstrap();
