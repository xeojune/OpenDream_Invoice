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

  await app.listen(8000);
}
bootstrap();
