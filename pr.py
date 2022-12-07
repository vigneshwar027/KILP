 where p2.PetitionerXref = '625365045' and
                
                ((DATEDIFF(YEAR,c.SourceCreatedDate,GETDATE())<=2) or
                (DATEDIFF(YEAR,c.CaseFiledDate,GETDATE())<=2) or
                (c.PrimaryCaseStatus = 'open' and b.IsActive=1) or
                
				(c.FinalAction in ('granted','denied') and
                (DATEDIFF(YEAR,c.FinalActionDate,GETDATE())<=1))) 