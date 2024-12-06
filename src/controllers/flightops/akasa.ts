import { Request, Response } from "express";
import { akasaPnrRetrieveUrl, akasaTokenUrl, environment } from "../../constants";
import moment from "moment-timezone";
import https from "https";
import xlsx from "xlsx";
import bluebird from "bluebird";
import axios, { AxiosResponse } from "axios";


// Set the default timezone for Moment.js
moment.tz.setDefault("Asia/Kolkata");

const AkasaTokenConfig = {
  method: "post",
  url: akasaTokenUrl,
  headers: {
    "Content-Type": "application/json",
    "User-Agent":
      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
  },
  data: JSON.stringify({ clientType: "WEB" }),
  httpsAgent: new https.Agent({ rejectUnauthorized: false }), // Bypass SSL verification
};

async function makeApiCall(
  maxAttempts: number,
  currentAttempt = 1
): Promise<AxiosResponse> {
  return axios(AkasaTokenConfig)
    .then((response) => {
      if (response.data.data.token && response.data.data.token.length > 0) {
        return response.data; // Return the response if a token is found
      } else if (currentAttempt < maxAttempts) {
        console.log(`Attempt ${currentAttempt} failed. Retrying...`);
        return makeApiCall(maxAttempts, currentAttempt + 1);
      } else {
        throw new Error("Maximum number of attempts reached");
      }
    })
    .catch((error) => {
      console.error(
        "Error fetching the token:",
        error.response ? error.response.data : error.message
      );
      if (currentAttempt < maxAttempts) {
        console.log(`Attempt ${currentAttempt} failed. Retrying...`);
        return makeApiCall(maxAttempts, currentAttempt + 1);
      } else {
        throw new Error("Maximum number of attempts reached after error");
      }
    });
}

const checkTimeFormat = (text: string): string => {
  if (
    text.charAt(0) === "0" ||
    text.slice(0, 2) === "11" ||
    text.slice(0, 2) === "10"
  ) {
    return text.concat(" ", "AM");
  } else {
    return text.concat(" ", "PM");
  }
};

export const getAkasaStatus = async (req:Request, res:Response) => {
    try {
         const data = await makeApiCall(10);
         const myToken = data?.data?.token;
         const header = myToken;

         const fullName = req.file?.filename;
         if (!fullName) {
           throw new Error("No file uploaded");
         }

         const wb = xlsx.readFile("./uploads/" + fullName, { cellDates: true });
         const ws = wb.Sheets["Sheet1"];
         ws["!ref"] = "A1:K3000"; // Adjust the range if necessary
         const jsonSheet = xlsx.utils.sheet_to_json(ws);
          const chunkSize = 30;
          const allResults = [];

           for (let i = 0; i < jsonSheet.length; i += chunkSize) {
             const chunk = jsonSheet.slice(i, i + chunkSize);
             const chunkResults = await Promise.all(
               chunk.map(async (record: any) => {
                 const PNR = record?.PNR;

                 // 3a. API requests with retry logic
                 const config1 = {
                   method: "get",
                   url: `${akasaPnrRetrieveUrl}?recordLocator=${PNR}&emailAddress=Airlines@Airiq.In`,
                   headers: {
                     Authorization: header,
                     "Content-Type": "application/json",
                     "User-Agent":
                       "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
                   },
                   httpsAgent: new https.Agent({ rejectUnauthorized: false }),
                 };
                //  const config2 = {
                //    method: "get",
                //    url: `${akasaPnrRetrieveUrl}?recordLocator=${PNR}&emailAddress=airlinesairiq@gmail.com`,
                //    headers: {
                //      Authorization: header,
                //      "Content-Type": "application/json",
                //      "User-Agent":
                //        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
                //    },
                //    httpsAgent: new https.Agent({ rejectUnauthorized: false }),
                //  };

                 let response: AxiosResponse<any>;
                 try {
                   response = await axios(config1);
                 } catch (error:any) {
                     return {
                       pnr: record.PNR,
                       error: `Error processing PNR ${PNR}: ${error?.message}`,
                   }
                 }

                 // 3b. Process API response
                 try {
                   if (!response || !response.data || !response.data.data) {
                     throw new Error("Invalid API response");
                   }

                   const bookingData = response.data.data;
                   const journeys = bookingData.journeys;
                   const length = journeys?.length;

                   if (length > 0) {
                     const PNR = bookingData?.recordLocator;
                     const Destination =
                       bookingData?.journeys[0].designator.destination;
                     const Origin = bookingData.journeys[0].designator.origin;
                     const depDate =
                       bookingData.journeys[0].designator.arrival.slice(0, 10);
                     const arrTime = checkTimeFormat(
                       bookingData.journeys[0].designator.arrival.slice(11, 16)
                     );
                     const depTime = checkTimeFormat(
                       bookingData.journeys[0].designator.departure.slice(
                         11,
                         16
                       )
                     );
                     const PAX = Object.keys(
                       bookingData.breakdown.passengers
                     ).length;
                     const flightReference =
                       bookingData.journeys[0].segments[0].flightReference;
                     const qpCodeRegex = /QP(\d+)/;
                     const match = flightReference.match(qpCodeRegex);
                     const flightNumber = match ? match[1] : null;

                     const OldPur = JSON.stringify(record.Pur);
                     const DepinputTime = record.Dep;
                     const date = moment.utc(DepinputTime).tz("Asia/Kolkata");
                     const minute = date.minutes();
                     const lastDigit = minute % 10;
                     let roundedMinutes = 0;
                     lastDigit === 0 || lastDigit === 5
                       ? (roundedMinutes = minute)
                       : (roundedMinutes = minute + 1);
                     date.minutes(roundedMinutes);
                     const OldDep = environment === "development" ? date.format("HH:mm A") : moment(date).subtract(5,"hours").subtract(30,"minutes");
                     const ArrinputTime = record.Arr;
                     const dateb = moment.utc(ArrinputTime).tz("Asia/Kolkata");
                     const minuteb = dateb.minutes();
                     const lastDigitb = minuteb % 10;
                     let roundedMinutesb = 0;
                     lastDigitb === 0 || lastDigitb === 5
                       ? (roundedMinutesb = minuteb)
                       : (roundedMinutesb = minuteb + 1);
                     dateb.minutes(roundedMinutesb);
                     const OldArr = environment === "development" ? dateb.format("HH:mm A") : moment(dateb).subtract(5,"hours").subtract(30,"minutes");;

                     const OldDate = moment(record.TravelDate).format(
                       "YYYY-MM-DD"
                     ); //Date
                     const CheckStatus = () => {
                       if (
                         record.Flight != flightNumber ||
                         OldPur != JSON.stringify(PAX) ||
                         OldDate != depDate ||
                         OldDep != depTime ||
                         OldArr != arrTime
                       ) {
                         return "BAD";
                       } else {
                         return "GOOD";
                       }
                     };

                     const MyRemarks = CheckStatus();
                     const result =
                       record.PNR +
                       "|" +
                       Origin +
                       "|" +
                       Destination +
                       "|" +
                       record.Flight +
                       "|" +
                       flightNumber +
                       "|" +
                       OldPur +
                       "|" +
                       PAX +
                       "|" +
                       OldDate +
                       "|" +
                       depDate +
                       "|" +
                       OldDep +
                       "|" +
                       depTime +
                       "|" +
                       OldArr +
                       "|" +
                       arrTime +
                       "|" +
                       MyRemarks;
                     return {
                       pnr: record.PNR,
                       data: result,
                     };
                   } else {
                     const PNR = bookingData.recordLocator;
                     return {
                       pnr: PNR,
                       data: `${PNR} is Cancelled`,
                     };
                   }
                 } catch (error: any) {
                   if (error?.response?.status === 404) {
                     return {
                       pnr: record.PNR,
                       error: `PNR NOT FOUND ${record.PNR}`,
                     };
                   } else {
                     return {
                       pnr: record.PNR,
                       error: `Error processing PNR ${record.PNR}: ${error?.message}`,
                     };
                   }
                 }
               })
             );
             allResults.push(...chunkResults);
           }
       const { results, errors } = allResults.reduce(
         (acc, result: any) => {
           if (result?.error) {
             acc.errors.push(result?.error);
           } else if (result.data) {
             acc.results.push(result.data); // Accumulate results as an array
           }
           return acc;
         },
         { results: [] as string[], errors: [] as string[] }
       );

       // Send the results as an array
       res.status(200).send({ results, errors });
    } catch (error) {
        res.status(500).send(error);
    }
};

// export const getAkasaStatus = async (req: Request, res: Response) => {
//   try {
//     const data = await makeApiCall(10);
//     const myToken = data?.data?.token;
//     const header = `Bearer ${myToken}`;

//     const fullName = req.file?.filename;
//     if (!fullName) {
//       throw new Error("No file uploaded");
//     }

//     const wb = xlsx.readFile("./uploads/" + fullName, { cellDates: true });
//     const ws = wb.Sheets["Sheet1"];
//     ws["!ref"] = "A1:K3000";
//     const jsonSheet = xlsx.utils.sheet_to_json(ws);

//     const chunkSize = 5;
//     const allResults: [] = [];

//     // Use Bluebird's map function with concurrency option
//     await bluebird.map(
//       jsonSheet,
//       async (record: any, index: number) => {
//         // Introduce a small delay to avoid overwhelming the API
//         await bluebird.delay(index * 100); // 100ms delay per request

//         const PNR = record?.PNR;

//         const config1 = {
//           method: "get",
//           url: `${akasaPnrRetrieveUrl}?recordLocator=${PNR}&emailAddress=Airlines@Airiq.In`,
//           headers: {
//             Authorization: header,
//             "Content-Type": "application/json",
//             "User-Agent":
//               "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
//           },
//           httpsAgent: new https.Agent({ rejectUnauthorized: false }),
//         };
//         const config2 = {
//           method: "get",
//           url: `${akasaPnrRetrieveUrl}?recordLocator=${PNR}&emailAddress=airlinesairiq@gmail.com`,
//           headers: {
//             Authorization: header,
//             "Content-Type": "application/json",
//             "User-Agent":
//               "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36",
//           },
//           httpsAgent: new https.Agent({ rejectUnauthorized: false }),
//         };

//         let response: AxiosResponse<any>;
//         try {
//           response = await axios(config1);
//         } catch (error) {
//           try {
//             response = await axios(config2);
//           } catch (error: any) {
//             return {
//               pnr: record.PNR,
//               error: `Error processing PNR ${PNR}: ${error?.message}`,
//             };
//           }
//         }

//         try {
//           if (!response || !response.data || !response.data.data) {
//             throw new Error("Invalid API response");
//           }

//           const bookingData = response.data.data;
//           const journeys = bookingData.journeys;
//           const length = journeys?.length;

//           if (length > 0) {
//             const Destination = bookingData?.journeys[0].designator.destination;
//             const Origin = bookingData.journeys[0].designator.origin;
//             const depDate = bookingData.journeys[0].designator.arrival.slice(
//               0,
//               10
//             );
//             const arrTime = checkTimeFormat(
//               bookingData.journeys[0].designator.arrival.slice(11, 16)
//             );
//             const depTime = checkTimeFormat(
//               bookingData.journeys[0].designator.departure.slice(11, 16)
//             );
//             const PAX = Object.keys(bookingData.breakdown.passengers).length;
//             const flightReference =
//               bookingData.journeys[0].segments[0].flightReference;
//             const qpCodeRegex = /QP(\d+)/;
//             const match = flightReference.match(qpCodeRegex);
//             const flightNumber = match ? match[1] : null;

//             const OldPur = JSON.stringify(record.Pur);
//             const DepinputTime = record.Dep;
//             const date = moment.utc(DepinputTime).tz("Asia/Kolkata");
//             const minute = date.minutes();
//             const lastDigit = minute % 10;
//             let roundedMinutes = 0;
//             lastDigit === 0 || lastDigit === 5
//               ? (roundedMinutes = minute)
//               : (roundedMinutes = minute + 1);
//             date.minutes(roundedMinutes);
//             const OldDep = date.format("HH:mm A");
//             const ArrinputTime = record.Arr;
//             const dateb = moment.utc(ArrinputTime).tz("Asia/Kolkata");
//             const minuteb = dateb.minutes();
//             const lastDigitb = minuteb % 10;
//             let roundedMinutesb = 0;
//             lastDigitb === 0 || lastDigitb === 5
//               ? (roundedMinutesb = minuteb)
//               : (roundedMinutesb = minuteb + 1);
//             dateb.minutes(roundedMinutesb);
//             const OldArr = dateb.format("HH:mm A");

//             const OldDate = moment(record.TravelDate).format("YYYY-MM-DD"); //Date
//             const CheckStatus = () => {
//               if (
//                 record.Flight != flightNumber ||
//                 OldPur != JSON.stringify(PAX) ||
//                 OldDate != depDate ||
//                 OldDep != depTime ||
//                 OldArr != arrTime
//               ) {
//                 return "BAD";
//               } else {
//                 return "GOOD";
//               }
//             };

//             const MyRemarks = CheckStatus();
//             const result =
//               record.PNR +
//               "|" +
//               Origin +
//               "|" +
//               Destination +
//               "|" +
//               record.Flight +
//               "|" +
//               flightNumber +
//               "|" +
//               OldPur +
//               "|" +
//               PAX +
//               "|" +
//               OldDate +
//               "|" +
//               depDate +
//               "|" +
//               OldDep +
//               "|" +
//               depTime +
//               "|" +
//               OldArr +
//               "|" +
//               arrTime +
//               "|" +
//               MyRemarks;
//             return {
//               pnr: record.PNR,
//               data: result,
//             };
//           } else {
//             const PNR = bookingData.recordLocator;
//             return {
//               pnr: PNR,
//               data: `${PNR} is Cancelled`,
//             };
//           }
//         } catch (error: any) {
//           if (error?.response?.status === 404) {
//             return {
//               pnr: record.PNR,
//               error: `PNR NOT FOUND ${record.PNR}`,
//             };
//           } else {
//             return {
//               pnr: record.PNR,
//               error: `Error processing PNR ${record.PNR}: ${error?.message}`,
//             };
//           }
//         }
//       },
//       { concurrency: chunkSize }
//     ); // Control concurrency with Bluebird

//     const { results, errors } = allResults.reduce(
//       (acc, result: any) => {
//         if (result?.error) {
//           acc.errors.push(result?.error);
//         } else if (result.data) {
//           acc.results.push(result.data); // Accumulate results as an array
//         }
//         return acc;
//       },
//       { results: [] as string[], errors: [] as string[] }
//     );

//     res.status(200).send({ results, errors });
//   } catch (error) {
//     res.status(500).send(error);
//   }
// };
