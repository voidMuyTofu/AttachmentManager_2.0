import { mergeStyleSets } from "@fluentui/react";

export const classNames = mergeStyleSets({
  fullWidthControl: {
    width: "100%",
  },
  fileIcon: {
    fontSize: 20,
  },
  wrapper: {
    height: "60vh",
    position: "relative",
  },
  filter: {
    paddingBottom: 20,
    maxWidth: 300,
  },
  header: {
    margin: 0,
  },
  row: {
    display: "inline-block",
  },
  buttonExaminar:{
    backgroundColor: "#E57E10",
    borderColor: "#E57E10",
    selectors: {
      ':hover': {
        backgroundColor: '#BF690D',
        borderColor: "#E57E10"
      },
      ':active': {
        backgroundColor: '#BF690D',
        borderColor: "#E57E10"
      }
    }
  },
});
