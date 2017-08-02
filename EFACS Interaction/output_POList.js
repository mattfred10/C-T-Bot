	var imports = JavaImporter();
	imports.importPackage(Packages.java.lang);
	imports.importPackage(Packages.uk.co.exel.base);
	imports.importPackage(Packages.uk.co.exel.bridge);
	imports.importPackage(Packages.uk.co.exel.framework.entitymanager);
	imports.importPackage(Packages.java.io.FileOutputStream);
	imports.importPackage(Packages.java.io.PrintStream);
	imports.importPackage(Packages.uk.co.exel.framework.entitymanager);
	var entityManager = imports.EntityManagerUtility.getEntityManager();
	



	// Set file location. We want to overwrite the previous file so that the PO List is always up to date.
	// Note that it has the filename 'EFACSPOList' to differentiate it from the POList saved by the python program.
	var POList = "Z:\\02 - Personal files\\Matt Frederick (Good)\\EFACSPOList.csv";
	out.println("## Collecting extant customer PO numbers ##");

	// Get the file stream and the printer
	var fout = new Packages.java.io.FileOutputStream(POList);
	var pout = new Packages.java.io.PrintStream(fout);

	var customers = ["vest01"]

	// The EQL query can only return lists of 1 column
	// Going to loop over customers
	for (customer in customers){
		// Set EQL query...
		var CustomerPOs = entityManager.createQuery("SELECT traderorderreference FROM SalesOrder WHERE traderid = :customer").setParameter("customer", customer).getResultList();
		// ...and turn it into an iterator					  
		var CustomerPOIterator = CustomerPOs.iterator();

		// 
		while (CustomerPOIterator.hasNext()){
			var PONum = CustomerPOIterator.next();
			pout.println(PONum);
		}
	}
		
	pout.close();
	fout.close();
	
	
} catch (error) {
	out.println(error);
}
