package com.xebia.documentSplitter.utils;

import java.io.File;
import java.io.IOException;

import org.apache.commons.io.FilenameUtils;

import com.appiancorp.suiteapi.common.exceptions.InvalidVersionException;
import com.appiancorp.suiteapi.common.exceptions.PrivilegeException;
import com.appiancorp.suiteapi.common.exceptions.StorageLimitException;
import com.appiancorp.suiteapi.content.ContentConstants;
import com.appiancorp.suiteapi.content.ContentService;
import com.appiancorp.suiteapi.content.exceptions.ContentExpiredException;
import com.appiancorp.suiteapi.content.exceptions.DuplicateUuidException;
import com.appiancorp.suiteapi.content.exceptions.InsufficientNameUniquenessException;
import com.appiancorp.suiteapi.content.exceptions.InvalidContentException;
import com.appiancorp.suiteapi.content.exceptions.NotLockOwnerException;
import com.appiancorp.suiteapi.content.exceptions.PendingApprovalException;
import com.appiancorp.suiteapi.knowledge.Document;

public class DocumentUtil {

    public static Long createDocument(ContentService cs, String name, String description, String extension, Long folder)
            throws InvalidContentException, PrivilegeException, InsufficientNameUniquenessException,
            StorageLimitException, DuplicateUuidException {

        Document doc = new Document();
        doc.setName(name);
        doc.setDescription(description);
        doc.setExtension(extension);
        doc.setParent(folder);
//        doc.setSize(size);

        return cs.create(doc, ContentConstants.UNIQUE_NONE);
    }

	public static Long updateDocument(ContentService cs, Long existingDocument, String name,
			String description/* , int size */)
            throws InvalidContentException, InvalidVersionException, PrivilegeException, NotLockOwnerException,
            InsufficientNameUniquenessException, PendingApprovalException, ContentExpiredException,
            StorageLimitException {
        Document doc = (Document) cs.getVersion(existingDocument, ContentConstants.VERSION_CURRENT);
        doc.setFileSystemId(ContentConstants.ALLOCATE_FSID);
        if (name != null && !name.isEmpty()) {
            doc.setName(name);
        }
        if (description != null && !description.isEmpty()) {
            doc.setDescription(description);
        }
//        doc.setSize(size);

        return cs.createVersion(doc, ContentConstants.UNIQUE_NONE).getId()[0];
    }

    public static File createFile(File file, String suffix, String extension) throws IOException {
        return createFile(FilenameUtils.getBaseName(file.getName()) + "_" + suffix, extension);
    }

    public static File createFile(String filename, String extension) throws IOException {
        return File.createTempFile(filename + "_", "." + extension);
    }
}
