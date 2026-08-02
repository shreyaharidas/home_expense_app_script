
export interface Colors {
  [key: string]: string;
}

export interface backgroundColors extends Colors {
  toEditColor: string;
  editedColor: string;
}
export const bgColors: backgroundColors = {
  toEditColor: "#FFBF00",
  editedColor: "#D1FFBD"
}