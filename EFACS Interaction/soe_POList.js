var importer = JavaImporter();
importer.importPackage(Packages.java.lang);
importer.importPackage(Packages.uk.co.exel.base);
importer.importPackage(Packages.uk.co.exel.bridge);
importer.importPackage(Packages.uk.co.exel.framework.entitymanager);

importer.importPackage(Packages.java.io.FileOutputStream);
importer.importPackage(Packages.java.io.PrintStream);

var esb = new importer.EmeraldSapphireBridge(this.appData.session());
var entityManager = importer.EntityManagerUtility.getEntityManager();


var soe = entityManager.get(“SalesOrder”, new Array(salesorderId));
var soCustRef = soe.getCustomerOrderReference();

var soCRIterator = soCustRef.iterator();

while (soCRIterator.hasNext()) {
	
	var PONum = soiIterator.next();
	
	
}