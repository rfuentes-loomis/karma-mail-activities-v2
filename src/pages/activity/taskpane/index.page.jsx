"use client";
//#region imports
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
import { handleAuth } from "../../../utils/ms-graph";
import Tooltip from "@mui/material/Tooltip";
import Menu from "@mui/material/Menu";
import IconButton from "@mui/material/IconButton";
import MoreVertIcon from "@mui/icons-material/MoreVert";
import Divider from "@mui/material/Divider";
import NotesIcon from "@mui/icons-material/Notes";
import AttachEmailIcon from "@mui/icons-material/AttachEmail";
import ArrowBackIosIcon from "@mui/icons-material/ArrowBackIos";
import Link from "@mui/material/Link";
import { getGraphData } from "../../../utils/ms-graph";
import Chip from "@mui/material/Chip";
import DeleteIcon from "@mui/icons-material/Delete";
import AddIcon from "@mui/icons-material/Add";
//#endregion

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

const insertActivityMutation = (payload) => applyMutation("addActivity", { input: payload }, ["workEffortId"]);

const postAttachments = ({ attachment, userId, workEffortId }) =>
  applyMutation("upsertActivityAttachment", {
    input: {
      attachmentSeqNumber: 1,
      contentBase64Encoded: attachment.contentBase64Encoded,
      fileName: attachment.fileName,
      userId: userId,
      workEffortId: workEffortId,
    },
  });

const fetchActiveEmployeeByEmail = (email) =>
  request(`{
    activeEmployeeByEmail(emailAddress:"${email}") {
      emailAddress
      ssiId
      userName
    }
  }`).then(({ activeEmployeeByEmail }) => activeEmployeeByEmail);

const getUniqueEmailBodyAsText = (msUser, messageId) =>
  getGraphData(msUser.auth.access_token, `/users/${msUser.id}/messages/${messageId}?$select=uniqueBody`, { Prefer: 'outlook.body-content-type="text"' });

const getEmailAsEml = (msUser, messageId) => getGraphData(msUser.auth.access_token, `/users/${msUser.id}/messages/${messageId}/$value`);

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
const useInsertActivity = () => useMutation({ mutationFn: (args) => insertActivityMutation(args) });

const useAuthUser = (enabled) => useQuery("user-auth", () => handleAuth(), { enabled, staleTime: 150 * 1000, refetchInterval: 150 * 1000 });

const useFetchActiveEmployeeByEmail = (email) => useQuery(`active-employee-${email}`, () => fetchActiveEmployeeByEmail(email), { enabled: email != null });

const useGetEmailUniqueBodyAsText = (currentMSUser, messageId) =>
  useQuery(`email-unique-body-${messageId}`, () => getUniqueEmailBodyAsText(currentMSUser, messageId), { enabled: false });

const useGetRawEmailContents = (currentMSUser, messageId) =>
  useQuery(`raw-email-${messageId}`, () => getEmailAsEml(currentMSUser, messageId), { enabled: false });

const usePostAttachments = () => useMutation({ mutationFn: (args) => postAttachments(args) });

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
  const [anchorEl, setAnchorEl] = useState(null);
  const open = Boolean(anchorEl);
  //#region Form config
  const formik = useFormik({
    initialValues: {
      userId: undefined,
      workEffortTypeId: undefined,
      workEffortCategoryId: "",
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
          userId: activeEmployee?.userName || currentMSUser?.surname,
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

        if (emailAsAttachment) {
          const attachmentPayload = { attachment: emailAsAttachment, userId: toSave.userId, workEffortId };
          try {
            const attachmentResults = await postAttachmentsAsync(attachmentPayload);
          } catch (error) {
            setFormErrorMsg("Activity Saved :),  However experienced error saving attachment.");
          }
        }

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
  const [emailItem, setEmailItem] = useState(null);
  const [emailAsAttachment, setEmailAsAttachment] = useState(null);

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
  const {
    data: workEffortCategories,
    isLoading: workEffortCategoriesLoading,
    isFetched: workEffortCategoriesIsFetched,
  } = useWorkEffortCategoryByEffortType(formik.values.workEffortTypeId);
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
  const { isFetching: fetchingEmailUniqueBody, refetch: fetchEmailUniqueBody } = useGetEmailUniqueBodyAsText(currentMSUser, emailItem?.itemId);
  const { isFetching: fetchingEmailRawContents, refetch: fetchEmailRawContents } = useGetRawEmailContents(currentMSUser, emailItem?.itemId);
  const { mutateAsync: postAttachmentsAsync, isSuccess: postAttachmentsSuccess } = usePostAttachments();
  //#endregion

  //#region Set default work Effort Category to Email
  useEffect(() => {
    if (workEffortCategoriesLoading || !workEffortCategoriesIsFetched) return;
    const emailType = workEffortCategories?.find((x) => toTrimAndLowerCase(x.workEffortCategoryDesc) === toTrimAndLowerCase("E-Mail"));
    formik.setFieldValue("workEffortCategoryId", emailType?.workEffortCategoryId);
  }, [workEffortCategories, workEffortCategoriesLoading, workEffortCategoriesIsFetched]);

  //#endregion

  //#region Handle Fatal Error
  useEffect(() => {
    if (activeEmployeeIsLoading || loadingMsUser) return;

    if (currentMsUserIsError) {
      setFatalError(true);
      setFatalErrorMsg(currentMsUserError?.name || "Error authenticating with Office.");
      return;
    }
    /**
     *
      Remove for now
     */
    // if (activeEmployeeIsError || (activeEmployeeIsFetched && activeEmployee === null)) {
    //   setFatalError(true);
    //   setFatalErrorMsg(`No record found for employee with email ${currentMSUser?.mail}`);
    //   return;
    // }
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

  //#region Handle Office On Ready & set Default values
  const officeOnReadyCallback = useCallback(() => {
    if (officeIsReady) return;
    if (!Office.context.mailbox) return;
    // Office.context.mailbox.item.to; //array [{emailAddress,displayName}]
    // Office.context.mailbox.item.bcc; //array [{emailAddress,displayName}]
    // Office.context.mailbox.item.cc; //array [{emailAddress,displayName}]
    // Office.context.mailbox.item.from; // object {emailAddress,displayName }
    setOfficeIsReady(true);
    setEmailItem(Office.context.mailbox.item);
    formik.setFieldValue("subject", Office.context.mailbox.item.normalizedSubject);
    formik.setFieldValue("workEffortDate", dayjs(new Date(Office.context.mailbox.item.dateTimeCreated)));
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
  const showLoading = () => workEffortTypesLoading || formik.isSubmitting || officeIsReady == false || loadingMsUser || fetchingEmailRawContents;
  const canSubmit = () => formik.isSubmitting === false && showLoading() === false;
  const handleOptionsMenuClick = (event) => setAnchorEl(event.currentTarget);
  const handleOptionsMenuClose = () => setAnchorEl(null);

  const copyEmailUniqueBodyToPublicComments = async () => {
    const { data } = await fetchEmailUniqueBody();
    formik.setFieldValue("commentsClob", data?.uniqueBody?.content);
  };

  const addEmailAsAttachment = async () => {
    const { data } = await fetchEmailRawContents();
    const contentBase64Encoded = btoa(data);
    const attachment = {
      contentBase64Encoded,
      fileName: `Email.eml`,
    };
    setEmailAsAttachment(attachment);
  };
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
            display: "flex",
            flexDirection: "column",
            height: "100vh",
            padding: "10px",
            "& .MuiFormControl-root": { margin: "7px 0px 7px 0px" },
          }}
        >
          {/* Header */}
          <Box
            sx={{
              display: "flex",
              flexDirection: "row",
              justifyContent: "space-between",
            }}
          >
            {/* Back button */}
            {formik.values.workEffortTypeId !== undefined ? (
              <IconButton
                size="small"
                sx={{ ml: 2 }}
                onClick={() => {
                  formik.setFieldValue("workEffortTypeId", undefined);
                  setEntitySearchStr("a");
                }}
              >
                <ArrowBackIosIcon />
              </IconButton>
            ) : null}

            {/* Title */}
            <h4 style={{ textAlign: "center", width: "100%" }}>
              {workEffortTypes?.find((x) => x.workEffortTypeId === formik.values.workEffortTypeId)?.subjectArea}
              Activity
            </h4>

            {/* Options Menu */}
            {formik.values.workEffortTypeId !== undefined ? (
              <>
                {" "}
                <Box sx={{ display: "flex", alignItems: "center", textAlign: "center" }}>
                  <Tooltip title="Email Options">
                    <IconButton
                      onClick={handleOptionsMenuClick}
                      size="small"
                      sx={{ ml: 2 }}
                      aria-controls={open ? "more-menu" : undefined}
                      aria-haspopup="true"
                      aria-expanded={open ? "true" : undefined}
                    >
                      <MoreVertIcon sx={{ width: 32, height: 32 }}></MoreVertIcon>
                    </IconButton>
                  </Tooltip>
                </Box>
                <Menu
                  anchorEl={anchorEl}
                  id="more-menu"
                  open={open}
                  onClose={handleOptionsMenuClose}
                  onClick={handleOptionsMenuClose}
                  PaperProps={{
                    elevation: 0,
                    sx: {
                      overflow: "visible",
                      filter: "drop-shadow(0px 2px 8px rgba(0,0,0,0.32))",
                      mt: 1.5,
                      "& .MuiSvgIcon-root": {
                        width: 32,
                        height: 32,
                        ml: -0.5,
                        mr: 2,
                      },
                      "&::before": {
                        content: '""',
                        display: "block",
                        position: "absolute",
                        top: 0,
                        right: 14,
                        width: 10,
                        height: 10,
                        bgcolor: "background.paper",
                        transform: "translateY(-50%) rotate(45deg)",
                        zIndex: 0,
                      },
                    },
                  }}
                  transformOrigin={{ horizontal: "right", vertical: "top" }}
                  anchorOrigin={{ horizontal: "right", vertical: "bottom" }}
                >
                  <MenuItem onClick={() => copyEmailUniqueBodyToPublicComments() && handleOptionsMenuClose()}>
                    <NotesIcon />
                    Copy email to comments
                    <Loading isLoading={fetchingEmailUniqueBody} />
                  </MenuItem>
                  <Divider />
                  <MenuItem onClick={() => addEmailAsAttachment() && handleOptionsMenuClose()}>
                    <AttachEmailIcon />
                    Add email thread as attachment
                  </MenuItem>
                </Menu>
              </>
            ) : null}
          </Box>
          {/* Sub header */}
          <Box>
            <Loading center isLoading={showLoading()} />
            {formErrorMsg ? (
              <Alert severity="error">
                <AlertTitle>Error</AlertTitle>
                {formErrorMsg}
              </Alert>
            ) : null}
          </Box>
          {/* Main Form Body */}
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
                    disabled={isWorkEffortAvailable(v.subjectArea) === false || showLoading()}
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
              <Box
                sx={{
                  display: "flex",
                  flexDirection: "column",
                }}
              >
                <TextField
                  id="commentsClob"
                  label="Public Comments"
                  multiline
                  rows={4}
                  helperText={formik.touched.commentsClob && formik.errors.commentsClob}
                  value={formik.values.commentsClob ?? ""}
                  onBlur={formik.handleBlur}
                  onChange={formik.handleChange}
                  error={formik.touched.commentsClob && Boolean(formik.errors.commentsClob)}
                />
                <Box
                  sx={{
                    display: "flex",
                    flexDirection: "row",
                    justifyContent: "space-between",
                  }}
                >
                  <Box>
                    <Loading isLoading={fetchingEmailUniqueBody} />
                  </Box>
                  <Link sx={{ float: "right", cursor: "pointer" }} onClick={copyEmailUniqueBodyToPublicComments}>
                    Copy email body
                  </Link>
                </Box>
              </Box>
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
            </>
          )}
          {/* Footer */}
          <Box>
            {formik.values.workEffortTypeId !== undefined ? (
              <>
                <Box
                  sx={{
                    display: "flex",
                    flexDirection: "row",
                  }}
                >
                  {emailAsAttachment ? (
                    <>
                      <Chip
                        sx={{
                          margin: "5px 0px 10px 0px",
                          "& .MuiChip-deleteIcon": {
                            color: "#bdbdbd",
                          },
                        }}
                        color="success"
                        variant="outlined"
                        avatar={<AttachEmailIcon />}
                        label={"Email added as attachment"}
                        onDelete={() => setEmailAsAttachment(null)}
                        deleteIcon={
                          <Tooltip title="Remove">
                            <DeleteIcon />
                          </Tooltip>
                        }
                      />
                    </>
                  ) : (
                    <>
                      <Chip
                        sx={{
                          margin: "5px 0px 10px 0px",
                          "& .MuiChip-deleteIcon": {
                            color: "#bdbdbd",
                          },
                        }}
                        variant="outlined"
                        avatar={<AttachEmailIcon />}
                        label={"Add email as attachment"}
                        onDelete={() => addEmailAsAttachment()}
                        deleteIcon={
                          <Tooltip title="Add email as attachment">
                            <AddIcon />
                          </Tooltip>
                        }
                      />
                    </>
                  )}
                  <Box
                    sx={{
                      margin: "10px 0px 0px 10px",
                    }}
                  >
                    <Loading isLoading={fetchingEmailRawContents} />
                  </Box>
                </Box>
                {/* Save Action */}

                <Button fullWidth variant="contained" type="submit" disabled={canSubmit() === false}>
                  {formik.isSubmitting ? <Loading isLoading /> : `Save`}
                </Button>
              </>
            ) : null}
          </Box>
        </Box>
      </form>
    </>
  );
};

export default App;
