import * as dotenv from "dotenv";

dotenv.config(); // Load environment variables from .env file

const environment = process.env.NODE_ENVIRONMENT;

const PORT = process.env.PORT || "";

const akasaPnrRetrieveUrl = process.env.AKASA_PNR_RETRIEVE_URL || " ";

const akasaTokenUrl = process.env.AKASA_TOKEN_URL || "";

const spicejetTokenUrl = process.env.SPICEJET_TOKEN_URL || "";

const spicejetPnrRetrieveUrl = process.env.SPICEJET_PNR_RETRIEVE_URL || "";



export {
    environment,
    PORT,
    akasaPnrRetrieveUrl,
    akasaTokenUrl,
    spicejetPnrRetrieveUrl,
    spicejetTokenUrl
}