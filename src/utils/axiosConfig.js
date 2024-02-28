import axios from "axios";

const api = axios.create({
  // eslint-disable-next-line no-undef
  baseURL: process.env.NEXT_PUBLIC_DHARMA_API,
});

// api.interceptors.response.use(
//   (response) => {
//     return response;
//   },
//   (error) => {
//     // we have to throw the error so react query goes into onError callbacks!
//     throw error;
//   }
// );

export { api };
