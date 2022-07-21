FROM node:18-alpine3.15

RUN mkdir -p /app/nfs/demandletters
RUN mkdir -p /app/docxletters
WORKDIR /app/docxletters

RUN chown node:node -R /app
USER node

# Install app dependency
COPY --chown=node package*.json ./
RUN npm install --production
COPY --chown=node . .

EXPOSE 8004

CMD ["node", "index.js"]

#  docker build -t migutak/docxletters:5.7.3 .

