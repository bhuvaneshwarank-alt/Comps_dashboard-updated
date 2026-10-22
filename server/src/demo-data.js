const masterHeaders=['Record ID','Record Title','Record Type','Department','Record Status','Primary Compliance Framework','Requirement / Control ID','Inherent Risk Rating','Residual Risk Rating','Owner','Due Date','Created Date','Target Completion Date','Action Status','Evidence Required?','Evidence Validated?','Description','Action Plan','Evidence Status'];
const records=[
 ['AUD-01-2026','Quarterly access review not documented','Audit Finding','IT','Open','SOC 2','CC6.2','High','High','IT Controls','2026-11-30','2026-10-03','2026-11-30','Planned','No','Yes','Synthetic demo record for the migration preview.','Document review and retain approval.','Missing'],
 ['RSK-01-2026','Evidence expiry tracking needs improvement','Risk','Compliance','In Progress','ISO 27001:2022','A.5.33','Medium','Medium','Compliance','2026-11-18','2026-10-05','2026-11-18','Planned','Yes','No','Synthetic demo record for the migration preview.','Confirm evidence owners and validity dates.','In progress'],
 ['INC-01-2026','Sample alert review follow-up','Incident','HR','Open','HIPAA','164.308(a)(1)','Medium','Low','HR Operations','2026-11-14','2026-10-09','2026-11-14','Planned','No','Yes','Synthetic demo record. Contains no person or patient information.','Complete the internal review.','Pending review'],
 ['AUD-02-2026','Training acknowledgement sample','Audit Finding','HR','Open','SOC 2','CC1.4','Low','Low','People Operations','2026-12-01','2026-10-13','2026-12-01','Planned','No','Yes','Synthetic demo record for the migration preview.','Upload approved training acknowledgement.','Pending'],
 ['SEC-01-2026','Sample endpoint inventory check','Security Finding','Admin','Closed','ISO 27001:2022','A.8.9','Low','Low','IT Admin','2026-10-19','2026-10-01','2026-10-19','Completed','No','Yes','Synthetic demo record for the migration preview.','Recorded for demo display only.','Complete'],
 ['RSK-02-2026','Business continuity tabletop evidence','Risk','IT','Open','HIPAA','164.308(a)(7)','Critical','High','IT Continuity','2026-11-02','2026-10-16','2026-10-01','Planned','No','Yes','Synthetic demo record for the migration preview.','Schedule a tabletop and capture results.','Overdue']
];
const userHeaders=['Email','Display Name','Role','Active','Updated By','Updated At','Departments'];
const users=[['Email','Display Name','Role','Active','Updated By','Updated At','Departments'],['admin@demo.invalid','Demo Administrator','Admin',true,'','2026-10-22','[]'],['hr@demo.invalid','Demo HR User','Department User',true,'','2026-10-22','["HR"]']];
const credentials=[['Email','Salt','Password Hash','Updated By','Updated At','Password History JSON','Password Expires At','Failed Attempts','Locked Until','Last Login At']];
const history=[
 ['Timestamp','Record ID','Action','Field Name','Old Value','New Value','Changed By','Change Reason'],
 ['2026-10-03T09:00:00.000Z','AUD-01-2026','CREATE','Record','', 'Quarterly access review not documented','migration-demo@invalid','Synthetic fixture'],
 ['2026-10-11T13:40:00.000Z','AUD-01-2026','STATUS UPDATE','Record Status','Open','Open','migration-demo@invalid','Synthetic fixture'],
 ['2026-10-09T08:20:00.000Z','INC-01-2026','CREATE','Record','','Sample alert review follow-up','migration-demo@invalid','Synthetic fixture']
];
export const demoConfig={masterHeaders,users,history,records};
export function demoTables(){return {
 User_Access:users,Role_Definitions:null,Temporary_Access_Grants:[['Grant ID','Request ID','User Email','Original Role','Original Departments JSON','Original Active','Granted Role','Granted Departments JSON','Start At','End At','Status','Created At','Activated At','Reverted At','Last Error']],
 Dashboard_Credentials:credentials,Compliance_Master_Register:[masterHeaders,...records],
 Record_Deletion_Register:[['Record ID','Deletion Status']],Record_History:history
};}
