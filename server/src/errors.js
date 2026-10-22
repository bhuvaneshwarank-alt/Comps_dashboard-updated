export class AppError extends Error {
  constructor(status,code,message){super(message);this.status=status;this.code=code;}
}
export function fail(status,code,message){throw new AppError(status,code,message);}
export const check=(condition,status,code,message)=>{if(!condition)fail(status,code,message);};
