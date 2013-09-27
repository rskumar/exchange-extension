package org.exoplatform.extension.exchange.listener;

import java.net.URI;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TimeZone;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.ScheduledFuture;
import java.util.concurrent.TimeUnit;

import microsoft.exchange.webservices.data.EventType;
import microsoft.exchange.webservices.data.ExchangeCredentials;
import microsoft.exchange.webservices.data.ExchangeService;
import microsoft.exchange.webservices.data.ExchangeVersion;
import microsoft.exchange.webservices.data.FolderEvent;
import microsoft.exchange.webservices.data.FolderId;
import microsoft.exchange.webservices.data.GetEventsResults;
import microsoft.exchange.webservices.data.ItemEvent;
import microsoft.exchange.webservices.data.PullSubscription;
import microsoft.exchange.webservices.data.WebCredentials;

import org.exoplatform.calendar.service.Calendar;
import org.exoplatform.container.xml.InitParams;
import org.exoplatform.extension.exchange.service.CalendarStorageService;
import org.exoplatform.extension.exchange.service.ExchangeIntegrationService;
import org.exoplatform.services.log.ExoLogger;
import org.exoplatform.services.log.Log;
import org.exoplatform.services.organization.OrganizationService;
import org.exoplatform.services.security.Identity;
import org.exoplatform.services.security.IdentityRegistry;
import org.picocontainer.Startable;

/**
 * 
 * Service to listen to exchange events. Used to synchronize eXo User calendar
 * state with exchange User calendar in real time. This service is used by two
 * modules: LoginModule to register User subscription to exchange events and a
 * LogoutListener that will be used to
 * 
 * @author Boubaker KHANFIR
 * 
 * 
 */
public class ExchangeListenerService implements Startable {

  private final static Log LOG = ExoLogger.getLogger(ExchangeListenerService.class);

  private static final String EXCHANGE_SERVER_URL_PARAM_NAME = "exchange.ews.url";
  private static final String EXCHANGE_DOMAIN_PARAM_NAME = "exchange.domain";
  private static final String EXCHANGE_LISTENER_SCHEDULER_DELAY_NAME = "exchange.scheduler.delay";

  private static long threadIndex = 0;
  private static short diffTimeZone = 0;
  private static int schedulerDelayInSeconds = 0;

  private final ScheduledExecutorService scheduledExecutor = Executors.newScheduledThreadPool(10);
  private final ExecutorService threadExecutor = Executors.newFixedThreadPool(10);
  private final Map<String, ScheduledFuture<?>> futures = new HashMap<String, ScheduledFuture<?>>();

  private final String exchangeServerURL;
  private final String exchangeDomain;

  private final CalendarStorageService calendarStorageService;
  private final OrganizationService organizationService;
  private final IdentityRegistry identityRegistry;

  public ExchangeListenerService(OrganizationService organizationService, CalendarStorageService calendarStorageService, IdentityRegistry identityRegistry, InitParams params) {
    this.calendarStorageService = calendarStorageService;
    this.identityRegistry = identityRegistry;
    this.organizationService = organizationService;

    if (params.containsKey(EXCHANGE_SERVER_URL_PARAM_NAME)) {
      exchangeServerURL = params.getValueParam(EXCHANGE_SERVER_URL_PARAM_NAME).getValue();
    } else {
      throw new IllegalStateException("Please add 'exchange.ews.url' parameter in configuration.properties.");
    }
    if (params.containsKey(EXCHANGE_DOMAIN_PARAM_NAME)) {
      exchangeDomain = params.getValueParam(EXCHANGE_DOMAIN_PARAM_NAME).getValue();
    } else {
      throw new IllegalStateException("Please add 'exchange.domain' parameter in configuration.properties.");
    }
    if (params.containsKey(EXCHANGE_LISTENER_SCHEDULER_DELAY_NAME)) {
      String schedulerDelayInSecondsString = params.getValueParam(EXCHANGE_LISTENER_SCHEDULER_DELAY_NAME).getValue();
      schedulerDelayInSeconds = Integer.valueOf(schedulerDelayInSecondsString);
    } else {
      throw new IllegalStateException("Please add 'exchange.scheduler.delay' parameter in configuration.properties.");
    }

    // Exchange system dates are saved using UTC timezone independing of User
    // Calendar timezone, so we have to get the diff with eXo Server TimeZone
    // and Exchange to make search queries
    diffTimeZone = getTimeZoneDiffWithUTC();

    LOG.info("Successfully started.");
  }

  @Override
  public void start() {
  }

  @Override
  public void stop() {
    scheduledExecutor.shutdownNow();
    threadExecutor.shutdownNow();
  }

  /**
   * Register User with Exchange services.
   * 
   * @param username
   * @param password
   */
  protected void userLoggedIn(final String username, final String password) {
    try {
      Identity identity = identityRegistry.getIdentity(username);
      if (identity == null) {
        throw new IllegalStateException("Identity of user '" + username + "' not found.");
      }

      // Scheduled task: listen the changes made on MS Exchange Calendar
      Thread schedulerCommand = new ExchangeIntegrationTask(username, password, false);
      ScheduledFuture<?> future = scheduledExecutor.scheduleWithFixedDelay(schedulerCommand, 30, schedulerDelayInSeconds, TimeUnit.SECONDS);

      // Execute the synchronization once logged in, but the first
      // synchronization have to be made after a subscription is done in the
      // previous scheduled thread, so that we can catch events even after
      // subscription
      Thread command = new ExchangeIntegrationTask(username, password, true);
      threadExecutor.execute(command);

      // Add future task to the map to destroy thread when the user logout
      {
        // Close other tasks if already exists, this can happens when user is
        // still logged in in other browser
        closeTaskIfExists(username);
        futures.put(username, future);
      }

      LOG.info("User '" + username + "' logged in, exchange synchronization task started.");
    } catch (Exception e) {
      LOG.warn("Error while initializing user '" + username + "' integration with exchange: " + e.getMessage());
      if (LOG.isTraceEnabled() || LOG.isDebugEnabled()) {
        LOG.trace("Error while initializing user integration with exchange: ", e);
      }
    }
  }

  /**
   * Unregister User from Exchange services.
   * 
   * @param username
   */
  protected void userLoggedOut(String username) {
    closeTaskIfExists(username);
  }

  private void closeTaskIfExists(String username) {
    ScheduledFuture<?> future = futures.get(username);
    if (future != null) {
      future.cancel(true);
      LOG.info("Exchange synchronization task stopped for User '" + username + "'.");
    }
  }

  private short getTimeZoneDiffWithUTC() {
    short diffTimeZone = 0;
    Date date = new Date();
    SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm");
    dateFormat.setTimeZone(TimeZone.getDefault());
    String dateTimeInOriginalTimeZone = dateFormat.format(date);
    dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
    String dateTimeInUTCTimeZone = dateFormat.format(date);

    try {
      long timeInOriginalTimeZone = dateFormat.parse(dateTimeInOriginalTimeZone).getTime();
      long timeInUTCTimeZone = dateFormat.parse(dateTimeInUTCTimeZone).getTime();
      diffTimeZone = (short) ((timeInUTCTimeZone - timeInOriginalTimeZone) / 3600000);
    } catch (Exception e) {
      LOG.error("Error while calculating difference between UTC Timezone and current one.");
    }

    return diffTimeZone;
  }

  /**
   * 
   * Thread used to synchronize Exchange Calendar with eXo Calendar
   * 
   */
  protected class ExchangeIntegrationTask extends Thread {
    private final ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
    private final ExchangeIntegrationService integrationService;
    private final List<FolderId> calendarFolderIds;
    private PullSubscription subscription = null;
    private String username;
    private boolean lazy;

    public ExchangeIntegrationTask(String username, String password, boolean lazy) throws Exception {
      super("ExchangeIntegrationTask-" + (threadIndex++));
      this.username = username;
      this.lazy = lazy;

      ExchangeCredentials credentials = new WebCredentials(username + "@" + exchangeDomain, password);
      service.setCredentials(credentials);
      service.setUrl(new URI(exchangeServerURL));

      integrationService = new ExchangeIntegrationService(organizationService, calendarStorageService, service, username);

      try {
        // First call to the service, this may fail because of wrong credentials
        calendarFolderIds = integrationService.getExchangeCalendars();
      } catch (Exception e) {
        if (lazy) {
          throw new RuntimeException("Error while authenticating user '" + username + "' to exchange, please make sure you are connected to the correct URL with correct credentials.", e);
        } else {
          throw e;
        }
      }
      if (!lazy) {
        // Don't need this subscription in case of first synchronization once
        // logged in
        newSubscription();
      }
    }

    @Override
    public void run() {
      try {
        long checkTime = java.util.Calendar.getInstance().getTimeInMillis();
        // This is used once, when user login
        if (lazy) {
          LOG.info("run first synchronization for user: " + username);
          Date lastSyncDate = integrationService.getUserLastCheckDate();

          // Verify if some folders was removed
          integrationService.synchronizeDeletedFolder(calendarFolderIds);

          // Verify modifications made on folders
          for (FolderId folderId : calendarFolderIds) {
            Calendar calendar = integrationService.getUserCalendar(folderId);
            if (calendar == null || lastSyncDate == null) {
              integrationService.synchronizeFullCalendar(folderId);
            } else {
              integrationService.synchronizeModificationsOfCalendar(folderId, lastSyncDate, diffTimeZone);
            }
          }
        } else {
          LOG.info("run scheduled synchronization for user: " + username);
          // This is used in a scheduled task when the user session still alive
          GetEventsResults events;
          try {
            events = subscription.getEvents();
          } catch (Exception e) {
            LOG.warn("Subscription seems timed out, retry. Original cause: " + e.getMessage() + "");

            newSubscription();
            events = subscription.getEvents();
          }

          Iterable<ItemEvent> itemEvents = events.getItemEvents();
          // If Calendar Folders was modified
          boolean folderCreatedOrDeleted = false;
          if (events.getFolderEvents() != null && events.getFolderEvents().iterator().hasNext()) {
            Iterator<FolderEvent> iterator = events.getFolderEvents().iterator();
            while (iterator.hasNext()) {
              FolderEvent folderEvent = (FolderEvent) iterator.next();
              if (folderEvent.getEventType().equals(EventType.Created) || folderEvent.getEventType().equals(EventType.Modified)) {
                if (!integrationService.isCalendarPresent(folderEvent.getFolderId())) {
                  boolean synchronizedCalendar = integrationService.synchronizeFullCalendar(folderEvent.getFolderId());
                  if (synchronizedCalendar && !calendarFolderIds.contains(folderEvent.getFolderId())) {
                    calendarFolderIds.add(folderEvent.getFolderId());
                  }
                  folderCreatedOrDeleted = true;
                }
              } else if (folderEvent.getEventType().equals(EventType.Deleted) || folderEvent.getEventType().equals(EventType.Moved)) {
                boolean deleted = integrationService.deleteCalendar(folderEvent.getFolderId());
                // If deleted, remove FolderId from listened folder Id and renew
                // subscription
                if (deleted && calendarFolderIds.contains(folderEvent.getFolderId())) {
                  calendarFolderIds.remove(folderEvent.getFolderId());
                  folderCreatedOrDeleted = true;
                }
              } else {
                LOG.warn("Folder Event wasn't catched: " + folderEvent.getEventType().name());
              }
            }
          }
          // loop through Appointment events
          for (ItemEvent itemEvent : itemEvents) {
            integrationService.createOrUpdateOrDelete(itemEvent);
          }

          // Renew subcription to manage events coming from new folder
          if (folderCreatedOrDeleted) {
            newSubscription();
          }
        }

        // Update date of last check in a user profile attribute
        integrationService.setUserLastCheckDate(checkTime);
      } catch (Exception e) {
        LOG.error("Error while synchronizing calndar entries.", e);
      }
    }

    @Override
    public void interrupt() {
      if (subscription != null) {
        try {
          LOG.info("Thread interruption: unsubscribe user service:" + username);
          subscription.unsubscribe();
        } catch (Exception e) {
          LOG.error("Thread interruption: Error while unsubscribe to thread of user:" + username);
        }
      }
      super.interrupt();
    }

    @Override
    protected void finalize() throws Throwable {
      if (subscription != null) {
        try {
          LOG.info("Object finalization: unsubscribe user service:" + username);
          subscription.unsubscribe();
        } catch (Exception e) {
          LOG.error("Object finalization: Error while unsubscribtion from exchange of user:" + username);
        }
      }
    }

    private void newSubscription() throws Exception {
      if (subscription != null) {
        try {
          subscription.unsubscribe();
        } catch (Exception e) {
          // Nothing to do, subscription may be timed out
          if (LOG.isDebugEnabled() || LOG.isTraceEnabled()) {
            LOG.error("Error while unsubscribe, will renew it anyway.", e);
          }
        }
      }
      subscription = service.subscribeToPullNotifications(calendarFolderIds, 5, null, EventType.Modified, EventType.Moved, EventType.FreeBusyChanged, EventType.Created, EventType.Deleted);
    }

  }
}
