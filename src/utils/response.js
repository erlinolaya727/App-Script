/**
 * Construye response para deterimnar resultado de un proceso
 * @param { ResponseCodes } status 
 * @param { Object }        body 
 * @returns { Object }
 */
const buildResponse = (status, body) => {
    return {
        status: status,
        body: body
    }
}