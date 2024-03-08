import axios from "axios";

const api = axios.create({
  // eslint-disable-next-line no-undef
  baseURL: process.env.NEXT_PUBLIC_DHARMA_API,
});


export { api };
