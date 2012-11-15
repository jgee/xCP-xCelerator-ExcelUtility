package com.emc.community.xcelerators.excelutil.impl;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

import com.documentum.fc.client.DfSingleDocbaseModule;
import com.documentum.fc.client.IDfSession;
import com.documentum.fc.client.IDfSysObject;
import com.documentum.fc.common.DfException;
import com.documentum.fc.common.DfId;
import com.documentum.fc.common.DfLogger;
import com.emc.community.xcelerators.excelutil.params.ColumnWithValues;
import com.emc.community.xcelerators.excelutil.params.CreatedDocumentId;
import com.emc.community.xcelerators.excelutil.params.NewObjectParams;
import com.emc.community.xcelerators.excelutil.params.ReadColumnWithValues;
import com.emc.community.xcelerators.excelutil.params.ReadingOptions;
import com.emc.community.xcelerators.excelutil.params.SheetReadRequest;
import com.emc.community.xcelerators.excelutil.params.StoreValuesParams;

/**
 * Implementation of the xCP ExcelUtil xCelerator
 * @author Michal.Malczewski
 *
 */
public class ExcelUtilImpl extends DfSingleDocbaseModule {

	// < > interface methods
	public CreatedDocumentId storeValuesAsCopy(StoreValuesParams manipulationParams, NewObjectParams newObjectParams) throws DfException {
		CreatedDocumentId result = new CreatedDocumentId();
		storeValues(manipulationParams, new SaveAsCopyPersister(result, newObjectParams));
		return result;
	}
	
	
	public void storeValuesInExistingSheet(StoreValuesParams params) throws DfException {
		storeValues(params, new SaveAsSameObjectPersister());
	}
	
	public ReadColumnWithValues[] readSheet(SheetReadRequest readRequest, ReadingOptions readingOptions) throws DfException {
		assertProvided(readRequest, "readRequest");
		String workbookId 	= getRequired(readRequest.getWorkbookObjectId(), "workbookObjectId");
		String sheetName	= getRequired(readRequest.getSheetName(), "sheetName");
		try {
			ExcelReader reader = new ExcelReader(getObjectContents(workbookId), readingOptions);
			return reader.readSheet(sheetName);
		} catch (DfException e) {
			throw e;
		} catch (Exception e) {
			throw new DfException(e);
		}
	}
	// </> interface methods
	
	// < > reading
	private InputStream getObjectContents(String objectId) throws DfException {
		IDfSession session = getSession();
		try {
			IDfSysObject excelSheet = (IDfSysObject) session.getObject(new DfId(objectId));
			return excelSheet.getContent();
		}
		finally {
			if (session != null) {
				releaseSession(session);
			}
		}	
	}
	// </> reading
	
	
	private void storeValues(StoreValuesParams params, ResultPersister persister) throws DfException {
		IDfSession session = getSession();
		try {
			try {
				storeValues(params, session, persister);
			} catch (DfException e) {
				throw e;
			} catch (Exception e) {
				throw new DfException(e);
			}
		}
		finally {
			releaseSession(session);
		}
	}
	
	protected void storeValues(StoreValuesParams params, IDfSession session, ResultPersister persister) throws Exception {
		assertProvided(params, "params");
		String sheetName 		= getRequired(params.getSheetName(), "sheetName");
		String sourceDocumentId	= getRequired(params.getSourceDocumentId(), "sourceDocumentId");
		ColumnWithValues[] rows	= getRequired(params.getColumns(), "columns");
		
		IDfSysObject excelSheet	= (IDfSysObject) session.getObject(new DfId(sourceDocumentId));
		InputStream originalConent = excelSheet.getContent();
		
		// temporary store result in memory
		ByteArrayOutputStream modifiedConetent = new ByteArrayOutputStream();		
		ExcelWriter writer = new ExcelWriter(originalConent);
		writer.storeValuesInSheet(sheetName, rows, params.getStartWritingInRow());
		writer.saveTo(modifiedConetent, params.isRecalculateFormulas());
		
		// save result in the repository
		persister.persist(modifiedConetent, excelSheet);
	}
	
	private <T> T getRequired(T value, String name) {
		assertProvided(value, name);
		return value;
	}


	private void assertProvided(Object value, String name) {
		if (value == null) {
			throw new IllegalStateException(name + " not provided");
		}
	}
	
	/**
	 * Saves the result
	 */
	private interface ResultPersister {
		
		/**
		 * Saves result according to the implementation
		 * @param modifiedContnent Result of the processing
		 * @param originalSheet Sysobject containing the original Excel sheet
		 */
		void persist(ByteArrayOutputStream modifiedConetent, IDfSysObject originalSheet) throws Exception;
	}
	
	
	private class SaveAsSameObjectPersister implements ResultPersister {

		public void persist(ByteArrayOutputStream modifiedConetent, IDfSysObject originalSheet) throws Exception {
			DfLogger.debug(this, "Saving results of object {0}", new Object[] {originalSheet.getObjectId()}, null);
			originalSheet.setContent(modifiedConetent);
			originalSheet.save();
		}
	}
	
	/**
	 * Saves result as a new object. If values in {@link NewObjectParams} passed are null, or it is null itself,
	 * an instance of dm_document will be created, named "{original sheet name} - filled" 
	 * @author Michal.Malczewski
	 *
	 */
	private class SaveAsCopyPersister implements ResultPersister {
		private final CreatedDocumentId result;
		private final NewObjectParams newObjectParams;
		
		public SaveAsCopyPersister(CreatedDocumentId result, NewObjectParams newObjectParams) {
			this.result = result;
			this.newObjectParams = newObjectParams;
		}

		public void persist(ByteArrayOutputStream modifiedConetent, IDfSysObject originalSheet) throws Exception {
			String objectType		= getValueOrReplacementIfNull((newObjectParams != null ? newObjectParams.getNewObjectType() 		: null), "dm_document");
			String primaryFolder	= getValueOrReplacementIfNull((newObjectParams != null ? newObjectParams.getNewObjectPrimaryFolder(): null), null);
			String objectName 		= getValueOrReplacementIfNull((newObjectParams != null ? newObjectParams.getNewObjectName() 		: null), originalSheet.getObjectName() + " - filled");
			
			IDfSysObject newObject = (IDfSysObject) originalSheet.getSession().newObject(objectType);
			newObject.setContentType(originalSheet.getContentType());
			newObject.setContent(modifiedConetent);
			newObject.setObjectName(objectName);
			if (primaryFolder != null) {
				newObject.link(primaryFolder);
			}
			newObject.save();
			this.result.setCreatedDocumentId(newObject.getObjectId().getId());
		}
		
		private String getValueOrReplacementIfNull(String value, String nullReplacement) {
			if (value == null) {
				return nullReplacement;
			}
			return value;
		}
	}
	
}
