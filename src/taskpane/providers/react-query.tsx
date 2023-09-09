import React from "react";
import { QueryClientProvider } from "@tanstack/react-query";
import { ReactFCC } from "../utils/ReactFCC";
import { queryClient } from "../lib/react-query";

export const ReactQueryProvider: ReactFCC = (props) => {
  // eslint-disable-next-line react/prop-types
  return <QueryClientProvider client={queryClient}>{props.children}</QueryClientProvider>;
};
