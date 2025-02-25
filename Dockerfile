# Use the official Node.js image.
FROM node:14

# Create and change to the app directory.
WORKDIR /usr/src/app

# Copy package.json and package-lock.json.
COPY package*.json ./

# Install app dependencies.
RUN npm install

# Copy the app source code.
COPY . .

# Expose the port the app runs on.
EXPOSE 3000

# Start the app.
CMD ["node", "src/app.js"]