package org.exoplatform.extension.exchange.service;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import javax.annotation.security.RolesAllowed;
import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.QueryParam;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;

import microsoft.exchange.webservices.data.Folder;
import microsoft.exchange.webservices.data.FolderId;

import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.rest.resource.ResourceContainer;
import org.exoplatform.services.security.ConversationState;

/**
 * 
 * @author Boubaker Khanfir
 * 
 */
@Path("/exchange")
public class ExchangeRESTService implements ResourceContainer, Serializable {
  private static final long serialVersionUID = -8085801604143848875L;

  private static final Log LOG = ExoLogger.getLogger(ExchangeRESTService.class);

  @GET
  @RolesAllowed("users")
  @Path("/calendars")
  @Produces({ MediaType.APPLICATION_JSON })
  public Response getCalendars() throws Exception {
    // It must be a user present in the session because of RolesAllowed
    // annotation
    String username = ConversationState.getCurrent().getIdentity().getUserId();

    List<FolderBean> beans = new ArrayList<FolderBean>();

    IntegrationService service = IntegrationService.getInstance(username);
    if (service != null) {
      List<FolderId> folderIDs = service.getAllExchangeCalendars();
      for (FolderId folderId : folderIDs) {
        Folder folder = service.getExchangeCalendar(folderId);
        if (folder != null) {
          boolean synchronizedFolder = service.isCalendarSynchronized(folderId.getUniqueId());
          FolderBean bean = new FolderBean(folderId.getUniqueId(), folder.getDisplayName(), synchronizedFolder);
          beans.add(bean);
        }
      }
    }
    return Response.ok(beans).build();
  }

  @GET
  @RolesAllowed("users")
  @Path("/sync")
  @Produces({ MediaType.APPLICATION_JSON })
  public Response synchronizeFolderWithExo(@QueryParam("folderId") String folderIdString) throws Exception {
    if (folderIdString == null || folderIdString.isEmpty()) {
      LOG.warn("folderId parameter is null while synchronizing.");
      return Response.noContent().build();
    }
    // It must be a user present in the session because of RolesAllowed
    // annotation
    String username = ConversationState.getCurrent().getIdentity().getUserId();
    IntegrationService service = IntegrationService.getInstance(username);
    service.addFolderToSynchronization(folderIdString);
    return Response.ok().build();
  }

  @GET
  @RolesAllowed("users")
  @Path("/unsync")
  @Produces({ MediaType.APPLICATION_JSON })
  public Response unsynchronizeFolderWithExo(@QueryParam("folderId") String folderIdString) throws Exception {
    if (folderIdString == null || folderIdString.isEmpty()) {
      LOG.warn("folderId parameter is null while unsynchronizing");
      return Response.noContent().build();
    }
    // It must be a user present in the session because of RolesAllowed
    // annotation
    String username = ConversationState.getCurrent().getIdentity().getUserId();
    IntegrationService service = IntegrationService.getInstance(username);
    service.deleteFolderFromSynchronization(folderIdString);
    return Response.ok().build();
  }

  public static class FolderBean implements Serializable {
    private static final long serialVersionUID = 4517749353533921356L;

    String id;
    String name;
    boolean synchronizedFolder = false;

    public FolderBean(String id, String name, boolean synchronizedFolder) {
      this.id = id;
      this.name = name;
      this.synchronizedFolder = synchronizedFolder;
    }

    public String getId() {
      return id;
    }

    public void setId(String id) {
      this.id = id;
    }

    public String getName() {
      return name;
    }

    public void setName(String name) {
      this.name = name;
    }

    public boolean isSynchronizedFolder() {
      return synchronizedFolder;
    }

    public void setSynchronizedFolder(boolean synchronizedFolder) {
      this.synchronizedFolder = synchronizedFolder;
    }
  }
}