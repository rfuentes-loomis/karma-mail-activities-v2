import { api } from "./axiosConfig";
import { EnumType, jsonToGraphQLQuery } from "json-to-graphql-query";

export const request = (query) => {
  return new Promise((resolve, reject) => {
    api
      .post("graphiql/query/", { query })
      .then((res) => {
        if (res.data.errors) {
          // console.error("QUERY error:", res.data.errors);
          // throw "Query syntax error";
          reject(res.data);
        }
        if (typeof res.data == "object") {
          // console.log(res.data.data);
          resolve(res.data.data);
        }
        reject(res.data);
        // else throw "Query error";
      })
      .catch((error) => {
        reject(error.response);
      });
  });
};

export const getMutation =
  (api, responseKeys = ["description", "status"]) =>
  (args) =>
    applyMutation(api, args, responseKeys);

export const applyMutation = (api, args, responseKeys = ["description", "status"]) => {
  const query = jsonToGraphQLQuery({
    mutation: {
      [api]: {
        __args: cleanPayload(args),
        ...createResponseMap(responseKeys),
      },
    },
  });
  // console.log("query", query);
  return request(query);
};

const createResponseMap = (responseKeys) => responseKeys.reduce((acc, key) => ({ ...acc, [key]: true }), {});

const cleanPayload = (payload) => {
  const ret = Object.entries(payload).reduce(
    (acc, [key, value]) => ({
      ...acc,
      [key]: value === undefined || value === "" ? null : isObject(value) && !(value instanceof EnumType) ? cleanPayload(value) : value,
    }),
    {}
  );
  // console.log("cleanedPayload", ret);
  return ret;
};

export const isObject = (variable) => typeof variable === "object" && !Array.isArray(variable) && variable != null;
