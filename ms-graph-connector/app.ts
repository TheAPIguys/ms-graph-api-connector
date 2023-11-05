import { APIGatewayProxyEvent, APIGatewayProxyResult } from 'aws-lambda'
import { processRequest } from './api-request/processRequest'

/**
 *
 * Event doc: https://docs.aws.amazon.com/apigateway/latest/developerguide/set-up-lambda-proxy-integrations.html#api-gateway-simple-proxy-for-lambda-input-format
 * @param {Object} event - API Gateway Lambda Proxy Input Format
 *
 * Return doc: https://docs.aws.amazon.com/apigateway/latest/developerguide/set-up-lambda-proxy-integrations.html
 * @returns {Object} object - API Gateway Lambda Proxy Output Format
 *
 */

export const lambdaHandler = async (event: APIGatewayProxyEvent): Promise<APIGatewayProxyResult> => {
  try {
    const bodyRequest = JSON.parse(event.body || '{}')
    console.log('Request Body: ', bodyRequest)
    const response = await processRequest(bodyRequest)
    return {
      statusCode: 200,
      body: JSON.stringify({
        count: response.length,
        data: response,
      }),
    }
  } catch (err) {
    console.log(err)
    return {
      statusCode: 500,
      body: JSON.stringify({
        message: err,
      }),
    }
  }
}
