import { createTheme } from "@mui/material/styles";

const Theme = createTheme({
  spacing: 2,
  palette: {
    primary: {
      light: "#759ab3",
      main: "#0c3b6d",
      dark: "#1f3c72", //"#001641",
    },
    secondary: {
      main: "#939848",
      light: "#ffb27d",
      dark: "#b95224",
    },
  },
  typography: {
    fontFamily: [
      "Source Sans Pro",
      "Open Sans",
      "Lato",
      "Roboto",
      "Avenir",
      '"Segoe UI"',
      "Arial",
      "Tahoma",
      "sans-serif",
    ].join(","),
    fontSize: 14,
  },
});

export default Theme;
