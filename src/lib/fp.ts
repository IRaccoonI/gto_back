type NotNullable<T> = T extends null | undefined ? never : T;

export const trusty = <T>(val: T): val is NotNullable<T> => {
  return val !== null && val !== undefined;
};

export const isNotEmptyObject = (val?: object | null): boolean => {
  return val !== null && val !== undefined && Object.keys(val).length !== 0;
};
