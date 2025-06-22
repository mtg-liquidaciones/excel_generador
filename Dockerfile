  FROM node:20-alpine
  WORKDIR /var/www/excel_generador
  COPY package*.json ./
  RUN npm ci --only=production
  COPY . .
  EXPOSE 9898
  USER node
  CMD [ "node", "/var/www/excel_generador/src/app.js" ]
