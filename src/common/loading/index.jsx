import * as React from "react";
import Box from "@mui/material/Box";
import CircularProgress from "@mui/material/CircularProgress";

// eslint-disable-next-line react/prop-types
const Loading = ({ isLoading, center }) =>
  isLoading ? (
    <Box sx={{ textAlign: center ? "center" : "inherit" }}>
      <CircularProgress size={20} />
    </Box>
  ) : null;

export default Loading;
