"use client";
import { useCallback, useState, useEffect } from "react";
import { request, applyMutation } from "../../../utils/graph-ql";
import { useMutation, useQuery } from "react-query";
import Box from "@mui/material/Box";
import TextField from "@mui/material/TextField";
import Button from "@mui/material/Button";
import InputLabel from "@mui/material/InputLabel";
import MenuItem from "@mui/material/MenuItem";
import FormControl from "@mui/material/FormControl";
import Select from "@mui/material/Select";
import { useFormik } from "formik";
import * as Yup from "yup";
import { DateTimePicker } from "@mui/x-date-pickers/DateTimePicker";
import dayjs from "dayjs";
import Autocomplete from "@mui/material/Autocomplete";
import Alert from "@mui/material/Alert";
import AlertTitle from "@mui/material/AlertTitle";
import Loading from "../../../common/loading";
import { toTrimAndLowerCase, extractNameFromEmail } from "../../../utils/stringHelpers";
import { handleAuth } from "../../../utils/auth";

//#region Graph QL query definition
const accountsSimpleGraph = `accounts{
  acid
  accountName
}`;
const clientSimpleGraph = `clients {
  ssiId
  clientName
}`;
const conferencesimpleGraph = ``;

const consultantsSimpleGraph = `consultants {
  ssiId
  name
}`;
const intermediarySimpleGraph = ``;

const investmentManagerSimpleGraph = `investmentManagers{
  ssiId
  name
}`;
const planSponsorSimpleGraph = `planSponsors {
  ssiId
  name
}`;
const privateClientSimpleGraph = `clients {
  ssiId
  clientName
}`;
const technologyClientSimpleGraph = ``;

const productsSimpleGraph = `products{
  ssiId
  productAlias
  productDesc
}`;
const employeesSimpleGraph = `employees {
  ssiId
  firstName
  lastName
}`;

const searchEntityQueryMap = [
  {
    entity: "account",
    query: accountsSimpleGraph,
    nameField: "accountName",
    idField: "acid",
  },
  {
    entity: "client",
    query: clientSimpleGraph,
    nameField: "clientName",
    idField: "ssiId",
  },
  {
    entity: "consultant",
    query: consultantsSimpleGraph,
    nameField: "name",
    idField: "ssiId",
  },
  {
    entity: "investmentManager",
    query: investmentManagerSimpleGraph,
    nameField: "name",
    idField: "ssiId",
  },
  {
    entity: "planSponsor",
    query: planSponsorSimpleGraph,
    nameField: "name",
    idField: "ssiId",
  },
  {
    entity: "product",
    query: productsSimpleGraph,
    nameField: "productDesc",
    idField: "ssiId",
  },
  {
    entity: "employee",
    query: employeesSimpleGraph,
    nameField: (v) => `${v.firstName} ${v.lastName}`,
    idField: "ssiId",
  },
];
//#endregion

//#region Api request
const searchDynamicEntity = (entityQueryMap, value) =>
  request(`{
  solrSearchText(searchText: "${value}"){
    ${entityQueryMap.query}
  }
}`).then(({ solrSearchText }) => solrSearchText[`${entityQueryMap.entity}s`]);

const fetchWorkEffortTypes = () =>
  request(`{
    workEffortTypes {
      subjectArea
      workEffortTypeDesc
      workEffortTypeId
    }
  }`).then(({ workEffortTypes }) => workEffortTypes);

const fetchWorkEffortCategoryByEffortType = (workEffortTypeId) =>
  request(`{
    workEffortCategories(workEffortTypeId: ${workEffortTypeId}) {
      workEffortCategoryDesc
      workEffortCategoryId
    }
  }`).then(({ workEffortCategories }) => workEffortCategories);

const insertActivityApi = (payload) => applyMutation("addActivity", { input: payload }, ["workEffortId"]);

const fetchActiveEmployeeByEmail = (email) =>
  request(`{
    activeEmployeeByEmail(emailAddress:"${email}") {
      emailAddress
      ssiId
      userName
    }
  }`).then(({ activeEmployeeByEmail }) => activeEmployeeByEmail);

//#endregion

//#region React Query Hooks
const useWorkEffortTypes = () => useQuery("WorkEffortTypes", fetchWorkEffortTypes);
const useWorkEffortCategoryByEffortType = (typeId) =>
  useQuery(`workCategoryByTypeId-${typeId}`, () => fetchWorkEffortCategoryByEffortType(typeId), { enabled: !!typeId });
const useSearchDynamicEntity = (entityQueryMap, value) => {
  return useQuery(`search-${entityQueryMap?.entity}-${value}`, () => searchDynamicEntity(entityQueryMap, value), {
    enabled: entityQueryMap != null,
    staleTime: Infinity,
  });
};
const useSearchProducts = (value) =>
  useQuery(`products-${value}`, () =>
    searchDynamicEntity(
      searchEntityQueryMap.find((x) => x.entity === "product"),
      value
    )
  );

const useSearchConsultants = (value) =>
  useQuery(`consultants-${value}`, () =>
    searchDynamicEntity(
      searchEntityQueryMap.find((x) => x.entity === "consultant"),
      value
    )
  );

const useSearchInvestmentManagers = (value) =>
  useQuery(`InvestmentManagers-${value}`, () =>
    searchDynamicEntity(
      searchEntityQueryMap.find((x) => x.entity === "investmentManager"),
      value
    )
  );

const useSearchEmployees = (value) =>
  useQuery(`Employees-${value}`, () =>
    searchDynamicEntity(
      searchEntityQueryMap.find((x) => x.entity === "employee"),
      value
    )
  );
const useInsertActivity = () => useMutation({ mutationFn: (args) => insertActivityApi(args) });

const useAuthUser = (enabled) => useQuery("user-auth", () => handleAuth(), { enabled: enabled });

const useFetchActiveEmployeeByEmail = (email) => useQuery(`active-employee-${email}`, () => fetchActiveEmployeeByEmail(email), { enabled: email != null });
//#endregion

//#region form schema
const ActivitySchema = Yup.object().shape({
  userId: Yup.string(),
  workEffortTypeId: Yup.number().required("Required"),
  workEffortCategoryId: Yup.number().required("Required"),
  workEffortDate: Yup.date().required("Required"),
  primaryItemId: Yup.number().required("Required"),
  subject: Yup.string(),
  commentsClob: Yup.string(),
  privateComments: Yup.string(),
  place: Yup.string(),
  planSponsors: Yup.array().of(
    Yup.object().shape({
      ssid: Yup.string(),
    })
  ),
  products: Yup.array().of(
    Yup.object().shape({
      ssid: Yup.string(),
    })
  ),
  representatives: Yup.array().of(
    Yup.object().shape({
      ssid: Yup.string(),
    })
  ),
});
//#endregion

const App = () => {
  //#region Form config
  const formik = useFormik({
    initialValues: {
      userId: undefined,
      workEffortTypeId: undefined,
      workEffortCategoryId: undefined,
      workEffortDate: dayjs(new Date()),
      primaryItemId: undefined,
      subject: undefined,
      place: undefined,
      commentsClob: undefined,
      privateComments: undefined,
      consultants: undefined,
      investmentManagers: undefined,
    },
    validationSchema: ActivitySchema,
    onSubmit: async (values, { setSubmitting: setFormikIsSubmitting }) => {
      setFormErrorMsg(null);
      setFormikIsSubmitting(true);
      try {
        const toSave = {
          ...values,
          userId: activeEmployee.userName,
          workEffortDate: values.workEffortDate.format(),
          consultants:
            values.consultants?.map((x) => ({
              ssiId: x.ssiId,
            })) || undefined,
          products:
            values.products?.map((x) => ({
              ssiId: x.ssiId,
            })) || undefined,
          representatives:
            values.representatives?.map((x) => ({
              ssiId: x.ssiId,
            })) || undefined,
          investmentManagers:
            values.investmentManagers?.map((x) => ({
              ssiId: x.ssiId,
            })) || undefined,
        };
        const results = await insertActivityAsync(toSave);
        const {
          addActivity: { workEffortId },
        } = results;

        window.scrollTo({ top: 0, behavior: "smooth" });
      } catch (ex) {
        setFormErrorMsg("Error saving activity.");
      }

      setSubmitting(false);
    },
  });
  //#endregion

  //#region State Variables
  const [fatalError, setFatalError] = useState(false);
  const [fatalErrorMsg, setFatalErrorMsg] = useState(null);
  const [officeIsReady, setOfficeIsReady] = useState(false);
  const [formErrorMsg, setFormErrorMsg] = useState(null);
  const [entitySearchStr, setEntitySearchStr] = useState("a");
  const [productsSearchStr, setProductsSearchStr] = useState("a");
  const [consultantsSearchStr, setConsultantsSearchStr] = useState("a");
  const [investmentManagersSearchStr, setInvestmentManagersSearchStr] = useState("a");
  const [employeesSearchStr, setEmployeesSearchStr] = useState("a");
  const [entityMapContext, setEntityMapContext] = useState(null);
  //#endregion

  //#region Queries & Mutations
  const { mutateAsync: insertActivityAsync, isSuccess: insertActivitySuccess } = useInsertActivity();
  const { data: workEffortTypes, isLoading: workEffortTypesLoading } = useWorkEffortTypes();
  const { data: workEffortCategories, isLoading: workEffortCategoriesLoading } = useWorkEffortCategoryByEffortType(formik.values.workEffortTypeId);
  const { data: dynamicSearchResults, isLoading: dynamicSearchResultsIsLoading } = useSearchDynamicEntity(entityMapContext, entitySearchStr);
  const { data: productsSearchResults, isLoading: productsSearchResultsIsLoading } = useSearchProducts(productsSearchStr);
  const { data: consultantsSearchResults, isLoading: consultantsSearchResultsIsLoading } = useSearchConsultants(consultantsSearchStr);
  const { data: investmentManagersSearchResults, isLoading: investmentManagersSearchResultsIsLoading } =
    useSearchInvestmentManagers(investmentManagersSearchStr);
  const { data: employeesSearchResults, isLoading: employeesSearchResultsIsLoading } = useSearchEmployees(employeesSearchStr);
  const { data: currentMSUser, isLoading: loadingMsUser, isError: currentMsUserIsError, error: currentMsUserError } = useAuthUser(officeIsReady);
  const {
    data: activeEmployee,
    isError: activeEmployeeIsError,
    isFetched: activeEmployeeIsFetched,
    isLoading: activeEmployeeIsLoading,
  } = useFetchActiveEmployeeByEmail(currentMSUser?.mail);
  //#endregion

  //#region Handle Fatal Error
  useEffect(() => {
    if (activeEmployeeIsLoading || loadingMsUser) return;

    if (currentMsUserIsError) {
      setFatalError(true);
      setFatalErrorMsg(currentMsUserError?.name || "Error authenticating with Office.");
      return;
    }

    if (activeEmployeeIsError || (activeEmployeeIsFetched && activeEmployee === null)) {
      setFatalError(true);
      setFatalErrorMsg(`No record found for employee with email ${currentMSUser?.mail}`);
      return;
    }
  }, [
    // current user
    loadingMsUser,
    currentMsUserIsError,
    currentMsUserError,
    currentMSUser,
    // active employee
    activeEmployeeIsLoading,
    activeEmployeeIsFetched,
    activeEmployeeIsError,
    activeEmployee,
  ]);
  //#endregion

  //#region Handle Office On Ready
  const officeOnReadyCallback = useCallback(() => {
    if (officeIsReady) return;
    setOfficeIsReady(true);
  }, [officeIsReady]);
  useEffect(() => {
    Office.onReady(officeOnReadyCallback);
  }, [officeOnReadyCallback]);
  useEffect(() => {
    if (!officeIsReady) return;
  }, [officeIsReady]);
  //#endregion

  //#region Helper Methods
  const isWorkEffortAvailable = (value) => toTrimAndLowerCase(value) === "plansponsor";
  const showLoading = () => workEffortTypesLoading || formik.isSubmitting || officeIsReady == false || loadingMsUser;
  const canSubmit = () => formik.isSubmitting === false && showLoading() === false;
  //#endregion

  if (fatalError) {
    return (
      <Box
        sx={{
          margin: "5px",
          display: "flex",
          flexDirection: "column",
          height: "100vh",
        }}
      >
        <Alert severity="error">
          <AlertTitle>Error. Please contact support</AlertTitle>
          {fatalErrorMsg}
        </Alert>
      </Box>
    );
  }

  if (insertActivitySuccess) {
    return (
      <Box
        sx={{
          margin: "5px",
          display: "flex",
          flexDirection: "column",
          height: "100vh",
        }}
      >
        <Alert severity="success">Activity Created.</Alert>
      </Box>
    );
  }

  return (
    <>
      <form onSubmit={formik.handleSubmit}>
        <Box
          sx={{
            margin: "5px",
            display: "flex",
            flexDirection: "column",
            height: "100vh",
            padding: "15px",
            "& .MuiFormControl-root": { margin: "5px 0px 5px 0px" },
          }}
        >
          <h3 style={{ textAlign: "center" }}>
            Create {workEffortTypes?.find((x) => x.workEffortTypeId === formik.values.workEffortTypeId)?.subjectArea}
            Activity
          </h3>
          <Loading center isLoading={showLoading()} />
          {formErrorMsg ? (
            <Alert severity="error">
              <AlertTitle>Error</AlertTitle>
              {formErrorMsg}
            </Alert>
          ) : null}
          {formik.values.workEffortTypeId === undefined ? (
            <>
              {/* 1st Step: workEffortType */}
              {workEffortTypes
                ?.sort((a, b) => {
                  if (a.subjectArea < b.subjectArea) {
                    return -1;
                  }
                  if (a.subjectArea > b.subjectArea) {
                    return 1;
                  }
                  return 0;
                })
                .map((v, k) => (
                  <Button
                    key={k}
                    sx={{ margin: "5px 0px 5px 0px" }}
                    variant="outlined"
                    disabled={isWorkEffortAvailable(v.subjectArea) === false}
                    onClick={() => {
                      formik.setFieldValue("workEffortTypeId", v.workEffortTypeId);
                      const entityMap = searchEntityQueryMap.find(
                        (x) =>
                          toTrimAndLowerCase(x.entity) ===
                          toTrimAndLowerCase(workEffortTypes?.find((x) => x.workEffortTypeId === v.workEffortTypeId)?.subjectArea)
                      );

                      setEntityMapContext(entityMap);
                    }}
                  >
                    {v.subjectArea}
                  </Button>
                ))}
              {/* <FormControl>
              <InputLabel variant="outlined" id="workEffortType-label">
                Entity
              </InputLabel>
              <Select
                disabled={workEffortTypesLoading}
                labelId="workEffortType-label"
                id="workEffortTypeId"
                label="Type"
                value={formik.values.workEffortTypeId}
                onBlur={formik.handleBlur}
                error={formik.touched.workEffortTypeId && Boolean(formik.errors.workEffortTypeId)}
                onChange={(event) => formik.setFieldValue("workEffortTypeId", event.target.value)}
              >
                {workEffortTypes?.workEffortTypes?.map((v, k) => (
                  <MenuItem key={k} value={v.workEffortTypeId}>
                    {v.subjectArea}
                  </MenuItem>
                ))}
              </Select>
            </FormControl> */}
            </>
          ) : (
            <>
              {/*  */}
              <Button
                variant="outlined"
                sx={{ margin: "10px 0px 10px 0px" }}
                onClick={() => {
                  formik.setFieldValue("workEffortTypeId", undefined);
                  setEntitySearchStr("a");
                }}
              >
                {`<-- Back`}
              </Button>
              {/* Choose Entity: primaryItemId */}
              <Autocomplete
                disablePortal
                value={formik.values.primaryItemId}
                id="primaryItemId"
                noOptionsText={"Type to search"}
                filterOptions={(x) => x}
                options={dynamicSearchResults || []}
                getOptionLabel={(option) => option[entityMapContext?.nameField]}
                getOptionKey={(option) => option[entityMapContext?.idField]}
                loading={dynamicSearchResultsIsLoading}
                onInputChange={(event, newValue) => setEntitySearchStr(newValue)}
                isOptionEqualToValue={(option, value) => option[entityMapContext?.idField] === value[entityMapContext?.idField]}
                onChange={(e, value) => formik.setFieldValue("primaryItemId", value[entityMapContext?.idField])}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    fullWidth
                    error={formik.touched.primaryItemId && Boolean(formik.errors.primaryItemId)}
                    helperText={formik.touched.primaryItemId && formik.errors.primaryItemId}
                    sx={{ marginRight: "5px" }}
                    label={workEffortTypes?.find((x) => x.workEffortTypeId === formik.values.workEffortTypeId)?.subjectArea}
                  />
                )}
              />
              {/* WorkEffortCategories */}
              <FormControl>
                <InputLabel variant="outlined" id="type-label">
                  Type
                </InputLabel>
                <Select
                  disabled={workEffortCategoriesLoading}
                  labelId="WorkEffortCategory-label"
                  id="workEffortCategoryId"
                  label="Type"
                  value={formik.values.workEffortCategoryId}
                  onBlur={formik.handleBlur}
                  helperText={formik.touched.workEffortCategoryId && formik.errors.workEffortCategoryId}
                  error={formik.touched.workEffortCategoryId && Boolean(formik.errors.workEffortCategoryId)}
                  onChange={(event) => formik.setFieldValue("workEffortCategoryId", event.target.value)}
                >
                  {workEffortCategories?.map((v, k) => (
                    <MenuItem key={k} value={v.workEffortCategoryId}>
                      {v.workEffortCategoryDesc}
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>
              {/* Subject */}
              <TextField
                id="subject"
                label="Subject"
                variant="outlined"
                helperText={formik.touched.subject && formik.errors.subject}
                value={formik.values.subject}
                onBlur={formik.handleBlur}
                onChange={formik.handleChange}
                error={formik.touched.subject && Boolean(formik.errors.subject)}
              />
              {/* Place */}
              <TextField
                id="place"
                label="Location"
                variant="outlined"
                value={formik.values.place}
                onBlur={formik.handleBlur}
                onChange={formik.handleChange}
                helperText={formik.touched.place && formik.errors.place}
                error={formik.touched.place && Boolean(formik.errors.place)}
              />
              {/* WorkEffortDate */}
              <DateTimePicker
                id="workEffortDate"
                value={formik.values.workEffortDate}
                error={formik.touched.workEffortDate && Boolean(formik.errors.workEffortDate)}
                helperText={formik.touched.workEffortDate && formik.errors.workEffortDate}
                onChange={(value) => formik.setFieldValue("workEffortDate", value)}
                referenceDate={dayjs(new Date())}
              />
              {/* Public Comments */}
              <TextField
                id="commentsClob"
                label="Public Comments"
                multiline
                rows={4}
                helperText={formik.touched.commentsClob && formik.errors.commentsClob}
                value={formik.values.commentsClob}
                onBlur={formik.handleBlur}
                onChange={formik.handleChange}
                error={formik.touched.commentsClob && Boolean(formik.errors.commentsClob)}
              />
              {/* Private Comments */}
              <TextField
                id="privateComments"
                label="Private Comments"
                multiline
                rows={4}
                helperText={formik.touched.privateComments && formik.errors.privateComments}
                value={formik.values.privateComments}
                onBlur={formik.handleBlur}
                onChange={formik.handleChange}
                error={formik.touched.privateComments && Boolean(formik.errors.privateComments)}
              />
              {/* Products */}
              <Autocomplete
                disablePortal
                multiple
                value={formik.values.products}
                id="products"
                noOptionsText={"Type to search"}
                filterOptions={(x) => x}
                options={productsSearchResults || []}
                getOptionLabel={(option) => option.productDesc}
                getOptionKey={(option) => option.ssiId}
                loading={productsSearchResultsIsLoading}
                onInputChange={(event, newValue) => setProductsSearchStr(newValue)}
                isOptionEqualToValue={(option, value) => option.ssiId === value.ssiId}
                onChange={(e, value) => formik.setFieldValue("products", value)}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    fullWidth
                    error={formik.touched.products && Boolean(formik.errors.products)}
                    helperText={formik.touched.products && formik.errors.products}
                    sx={{ marginRight: "5px" }}
                    label={"Products"}
                  />
                )}
              />
              {/* Consultants */}
              <Autocomplete
                disablePortal
                multiple
                value={formik.values.consultants}
                id="consultants"
                noOptionsText={"Type to search"}
                filterOptions={(x) => x}
                options={consultantsSearchResults || []}
                getOptionLabel={(option) => option.name}
                getOptionKey={(option) => option.ssiId}
                loading={consultantsSearchResultsIsLoading}
                onInputChange={(event, newValue) => setConsultantsSearchStr(newValue)}
                isOptionEqualToValue={(option, value) => option.ssiId === value.ssiId}
                onChange={(e, value) => formik.setFieldValue("consultants", value)}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    fullWidth
                    error={formik.touched.consultants && Boolean(formik.errors.consultants)}
                    helperText={formik.touched.consultants && formik.errors.consultants}
                    sx={{ marginRight: "5px" }}
                    label={"Consultants"}
                  />
                )}
              />
              {/* InvestmentManagers */}
              <Autocomplete
                disablePortal
                multiple
                value={formik.values.investmentManagers}
                id="investmentManagers"
                noOptionsText={"Type to search"}
                filterOptions={(x) => x}
                options={investmentManagersSearchResults || []}
                getOptionLabel={(option) => option.name}
                getOptionKey={(option) => option.ssiId}
                loading={investmentManagersSearchResultsIsLoading}
                onInputChange={(event, newValue) => setInvestmentManagersSearchStr(newValue)}
                isOptionEqualToValue={(option, value) => option.ssiId === value.ssiId}
                onChange={(e, value) => formik.setFieldValue("investmentManagers", value)}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    fullWidth
                    error={formik.touched.investmentManagers && Boolean(formik.errors.investmentManagers)}
                    helperText={formik.touched.investmentManagers && formik.errors.investmentManagers}
                    sx={{ marginRight: "5px" }}
                    label={"Investment Managers"}
                  />
                )}
              />
              {/* Loomis Representatives  */}
              <Autocomplete
                disablePortal
                multiple
                value={formik.values.investmentManagers}
                id="representatives"
                noOptionsText={"Type to search"}
                filterOptions={(x) => x}
                options={employeesSearchResults || []}
                getOptionLabel={(option) => `${option.firstName} ${option.lastName}`}
                getOptionKey={(option) => option.ssiId}
                loading={employeesSearchResultsIsLoading}
                onInputChange={(event, newValue) => setEmployeesSearchStr(newValue)}
                isOptionEqualToValue={(option, value) => option.ssiId === value.ssiId}
                onChange={(e, value) => formik.setFieldValue("representatives", value)}
                renderInput={(params) => (
                  <TextField
                    {...params}
                    fullWidth
                    error={formik.touched.representatives && Boolean(formik.errors.representatives)}
                    helperText={formik.touched.representatives && formik.errors.representatives}
                    sx={{ marginRight: "5px" }}
                    label={"Loomis Representatives"}
                  />
                )}
              />
              {/* Save Action */}
              <Button fullWidth variant="contained" type="submit" disabled={canSubmit() === false}>
                {formik.isSubmitting ? <Loading isLoading /> : `Save`}
              </Button>
            </>
          )}
        </Box>
      </form>
    </>
  );
};

export default App;
