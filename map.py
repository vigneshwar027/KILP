headers = ['Beneficiary Id', 'Case No.', 'Beneficiary Type', 'Beneficiary Record Creation Date', 'Beneficiary Status', 'Organization', 'Petitioner', 'Petitioner of Primary Beneficiary', 'Beneficiary Last Name', 'Beneficiary First Name', 'Primary Beneficiary Id', 'Primary Beneficiary Last Name', 'Primary Beneficiary First Name', 'Relation', 'Immigration Status', 'Immigration Status Exp Date', 'Process Opened Date', 'Process Id', 'Process Type', 'Process Reference', 'Process Filed Date', 'Primary Process Status', 'Secondary Process Status', 'Secondary Process Status Date', 'Final Action', 'Final Action Date', 'Case Closed', 'Employee Id', 'Manager Name', 'Business Partner Name', 'Dept', 'Dept Group', 'Dept Number', 'Cost Center', 'Cost Ctr No. ', 'Business Unit Code', 'VP', 'TPX Project', 'Questionnaire Sent to Manager', 'Questionnaire Sent to FN', 'Questionnaire Returned by Manager', 'Questionnaire Returned by FN', 'All Petitioning Company Info Received', 'All FN Docs Received', 'LCA Filed', 'Documents Sent for Signature', 'Singed Docs Received', 'AOS Docs Sent for Signature', 'Signed AOS Docs Received', 'RFE Received', 'RFE Due Date', 'RFE Docs Requested', 'RFE Docs Received', 'RFE Docs to ER for Review / Signature', 'RFE Response Submitted', 'PERM Memo Sent to Employer', 'Approval of PERM Memo Received', 'Employee Work Experience Chart Sent', 'Employee Work Experience Chart Received', 'PWD Request Submitted to DOL', 'PWD Issued by DOL', 'PWD Expiration Date', 'Recruitment Approval Received from ER', 'Recruitment Instructions Sent to Company', 'Job Order Placed with SWA', 'Dated Copies of All Recruitment Received', 'Recruitment Report Sent to Company', 'Recruitment Report Received', 'Form 9089 Sent to FN and Employer', 'Edits to Form 9089 Received from FN and Employer', 'Form 9089 Submitted to DOL', 'Audit Notice Received', 'Audit Docs to ER for Review / Signature', 'Audit Docs Received from ER', 'Audit Response Sent to DOL']



#headers as in db
headers_table = ['BeneficiaryXref', 'Beneficiary_Xref2', 'BeneficiaryType', 'SourceCreatedDate', 'IsActive', 'OrganizationName', 'PetitionerName', 'PetitionerOfPrimaryBeneficiary', 'LastName', 'FirstName', 'PrimaryBeneficiaryXref', 'PrimaryBeneficiaryLastName', 'PrimaryBeneficiaryFirstName', 'RelationType', 'Current_Immigration_Status', 'CurrentImmigrationStatusExpirationDate2', 'SourceCreatedDate', 'CaseXref', 'CaseType', 'CaseDescription', 'CaseFiledDate', 'PrimaryCaseStatus', 'LastStepCompleted', 'LastStepCompletedDate', 'FinalAction', 'FinalActionDate', 'CaseClosedDate', 'EmployeeId', 'ManagerName', 'BusinessPartnerName', 'Department', 'Department_Group', 'Department_Number', 'CostCenter', 'CostCenterNumber', 'BusinessUnitCode', 'SecondLevelManager', 'TPX_PROJECT', 'questionnairesenttomanager', 'questionnairessenttofn', 'questionnairecompletedandreturnedbymanager', 'questionnairecompletedandreturnedbyfn', 'allpetitioningcompanyinforeceived', 'allfndocsreceived', 'lcafiled', 'formsanddocumentationsubmittedforsignature', 'signedformsandletterreceived', 'dateaosformssentforsignature', 'datesignedaosformsreceived', 'RFEAuditReceivedDate', 'RFEAuditDueDate', 'RFEDocsReqestedDate', 'RFEDocsReceivedDate', 'RFE_Docs_to_ER_for_Review_Signature', 'RFEAuditSubmittedDate', 'permmemosenttoemployer', 'approvalofpermmemoreceived', 'employeeworkexperiencechartsent', 'employeeworkexperiencechartreceived', 'prevailingwagedeterminationrequestsubmittedtodol', 'prevailingwagedeterminationissuedbydol', 'PwdExpirationDate', 'Recruitment_Approval_Received_from_ER', 'recruitmentinstructionssenttocompany', 'joborderplacedwithswa', 'datedcopiesofallrecruitmentreceived', 'recruitmentreportsenttocompany', 'recruitmentreportreceived', 'form9089senttofnandemployer', 'editstoform9089receivedfromfnandemployer', 'form9089submittedtodol', 'PERMAuditReceivedDate', 'Audit_Docs_to_ER_for_Review_Signature', 'Audit_Docs_Received_from_ER', 'PERMAuditSubmittedDate']


date_columns = ['SourceCreatedDate', 'CurrentImmigrationStatusExpirationDate2', 'SourceCreatedDate', 'CaseFiledDate', 'LastStepCompletedDate', 'FinalActionDate', 'CaseClosedDate', 'questionnairessenttofn', 'questionnairesenttomanager', 'questionnairecompletedandreturnedbymanager', 'questionnairecompletedandreturnedbyfn', 'allpetitioningcompanyinforeceived', 'allfndocsreceived', 'lcafiled', 'formsanddocumentationsubmittedforsignature', 'signedformsandletterreceived', 'dateaosformssentforsignature', 'datesignedaosformsreceived', 'RFEAuditReceivedDate','RFE_Docs_to_ER', 'RFEAuditDueDate', 'RFEDocsReqestedDate', 'RFEDocsReceivedDate', 'RFEAuditSubmittedDate', 'permmemosenttoemployer', 'approvalofpermmemoreceived', 'employeeworkexperiencechartsent', 'employeeworkexperiencechartreceived', 'prevailingwagedeterminationrequestsubmittedtodol', 'prevailingwagedeterminationissuedbydol', 'PwdExpirationDate', 'recruitmentinstructionssenttocompany', 'joborderplacedwithswa', 'datedcopiesofallrecruitmentreceived', 'recruitmentreportsenttocompany', 'recruitmentreportreceived', 'form9089senttofnandemployer', 'editstoform9089receivedfromfnandemployer', 'form9089submittedtodol', 'PERMAuditReceivedDate', 'PERMAuditSubmittedDate']


# a = a.split('\n')
# # a = ','.join(a)
# # print(a)

# with open('op.txt','w') as file:
# 	file.write(str(a))


for i in headers:
    print(i)
print('\n\n\n')
for i in headers_table:
    print(i)


	# (c.PrimaryCaseStatus = 'open' and c.CasePetitionId in ('100003008','100003034','100003010','100003009','100003013')) or
		
	#       (c.CasePetitionId = '100003008' and c.PrimaryCaseStatus = 'closed' and  (datediff(year,c.CaseOpenDate,getdate()) <1)) or

	# 	  (c.CasePetitionId = '100003034' and
	# 	   c.CaseDescription in ('Change of Employer','COE', 'H-1B Change COE', 'H-1B Change of Employer', 'H-1B Change of ER' ,'Ext', 'Extension', 'H-1B Ext', 'H-1B Extension') and
	# 	   c.PrimaryCaseStatus = 'closed' and
	# 	   (datediff(year,c.CaseOpenDate,getdate()) <1)) or

	# 	   (c.CasePetitionId = '100003010' and
	# 	   c.PrimaryCaseStatus = 'closed' and
	# 	   (datediff(year,c.CaseOpenDate,getdate()) <1)) or

	# 	   (c.CasePetitionId = '100003009' and
	# 	   c.PrimaryCaseStatus = 'closed' and
	# 	   (datediff(year,c.CaseOpenDate,getdate()) <1)) or	

	# 	   (c.CasePetitionId = '100003013' and
	# 	   c.PrimaryCaseStatus = 'closed' and
	# 	   (datediff(year,c.CaseOpenDate,getdate()) <1)) or	

	# 	   (c.CaseType != 'Labor Cert PERM' and
	# 	   (datediff(year,c.RFEAuditReceivedDate,getdate()) <1)) or	

	# 	   (c.CaseType = 'Labor Cert PERM' and
	# 	   (datediff(year,c.PERMAuditReceivedDate,getdate()) <1))	