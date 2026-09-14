package org.joget.mokxa;

import org.joget.apps.app.service.AppUtil;
import org.joget.apps.datalist.model.*;
import org.joget.apps.form.model.FormRow;
import org.joget.apps.form.model.FormRowSet;
import org.joget.commons.util.LogUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.http.client.fluent.Content;
import org.apache.http.client.fluent.Form;
import org.apache.http.client.fluent.Request;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;

import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * OutlookEmailListBinder – Optimised with Server‑Side Date Filters + Client‑Side Subject Filter
 *
 * Lists primary (sent) emails tagged with a Matter ID extended property.
 * Supports optional server‑side date range filtering using two hidden DataList filter columns:
 *   - start_filter   → sentDateTime ge
 *   - end_filter     → sentDateTime le
 * Subject filtering is applied locally after the Graph response is received.
 * Uses the Sent Items folder to guarantee no duplicate messages.
 * Efficient Graph‑level pagination and per‑thread unread counts.
 */
public class OutlookEmailListBinder extends DataListBinderDefault {

    private static final String TAG = "OutlookEmailListBinder";
    private static final String GRAPH_BASE = "https://graph.microsoft.com/v1.0/users/";
    private static final String TOKEN_BASE = "https://login.microsoftonline.com/";
    private static final String GRAPH_SCOPE = "https://graph.microsoft.com/.default";
    private static final ObjectMapper MAPPER = new ObjectMapper();
    private static final DateTimeFormatter DISPLAY_FMT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    // ----------------------------------------------------------------
    // Plugin metadata
    // ----------------------------------------------------------------
    @Override public String getName()        { return getClass().getName(); }
    @Override public String getVersion()     { return "3.0.0"; }
    @Override public String getLabel()       { return "Outlook Sent Email List (Primary, Date+Subject)"; }
    @Override public String getDescription() {
        return "Lists sent primary emails tagged with Matter ID. " +
                "Supports server‑side date filters (start_filter / end_filter). " +
                "Subject filtering is applied after fetching (client‑side). " +
                "Efficient pagination, unread count per thread.";
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(),
                "/properties/OutlookEmailListBinder.json", null, true, null);
    }

    // ----------------------------------------------------------------
    // Columns
    // ----------------------------------------------------------------
    @Override
    public DataListColumn[] getColumns() {
        return new DataListColumn[]{
                new DataListColumn("id",                "Message ID",          false),
                new DataListColumn("subject",           "Subject",             true),
                new DataListColumn("from",              "From",                true),
                new DataListColumn("toRecipients",      "To",                  true),
                new DataListColumn("dateCreated",       "Date Created",        true),
                new DataListColumn("unreadCount",       "Unread Count",        true),
                new DataListColumn("hasAttachments",    "Attachments",         true),
                new DataListColumn("conversationId",    "Conversation ID",     false),
                new DataListColumn("internetMessageId", "Internet Message ID", false),
                new DataListColumn("matterId",          "Matter ID",           true),
                new DataListColumn("start_filter",      "Start Filter",           true),
                new DataListColumn("end_filter",        "End Filter",           true),
                new DataListColumn("internetMessageIdEncoded", "Internet Message ID Encoded", false),


        };
    }

    @Override public String getPrimaryKeyColumnName() { return "id"; }

    @Override
    public DataListCollection getData(DataList dataList, Map properties,
                                      DataListFilterQueryObject[] filters,
                                      String sort, Boolean desc,
                                      Integer start, Integer rows) {

        DataListCollection result = new DataListCollection();

        int pageSize = (rows  != null && rows  > 0) ? rows  : 10;
        int skip     = (start != null && start > 0) ? start : 0;

        try {
            String token = getAppOnlyAccessToken();
            if (token == null) return result;

            String extPropId     = buildExtendedPropertyId();
            String matterIdValue = getPropertyString("matterIdValue");
            String principalUser = getPropertyString("principalUserEmail");

            if (extPropId == null || StringUtils.isBlank(matterIdValue) ||
                    StringUtils.isBlank(principalUser)) {
                LogUtil.warn(TAG, "Missing required plugin properties.");
                return result;
            }

            // Extract optional filters
            String dateFrom = null;
            String dateTo   = null;
            String subject  = null;

            if (filters != null) {
                for (DataListFilterQueryObject f : filters) {
                    String q = (f.getQuery() != null) ? f.getQuery().toLowerCase() : "";
                    String[] vals = f.getValues();
                    if (vals == null || vals.length == 0) continue;
                    String v0 = vals[0];
                    if (StringUtils.isBlank(v0)) continue;

                    if (q.contains("subject")) {
                        subject = v0;
                    } else if (q.contains("start_filter")) {
                        dateFrom = toUtcIso(v0, false);
                    } else if (q.contains("end_filter")) {
                        dateTo   = toUtcIso(v0, true);
                    }
                }
            }

            // Build main query (Sent Items folder, extended property, optional dates)
            // Subject is NOT sent to Graph
            String url = buildPrimaryUrl(principalUser, extPropId, matterIdValue,
                    dateFrom, dateTo,subject, pageSize, skip);
//            LogUtil.info(TAG, "[getData] URL: " + url);

            JsonNode messages = fetchJsonArray(url, token);
            if (messages == null) return result;
            FormRowSet rowSet = new FormRowSet();

            for (JsonNode msg : messages) {
                String convId = text(msg, "conversationId");
                int unread = 0;

                if (StringUtils.isNotBlank(convId)) {
                    unread = fetchUnreadCount(principalUser, convId, token);
                }

                rowSet.add(toFormRow(msg, unread));
            }

            result.addAll(rowSet);

        } catch (Exception e) {
            LogUtil.error(TAG, e, "[getData] Error");
        }
        return result;
    }


    @Override
    public int getDataTotalRowCount(DataList dataList, Map properties,
                                    DataListFilterQueryObject[] filters) {
        try {
            String token = getAppOnlyAccessToken();
            if (token == null) return 0;

            String extPropId     = buildExtendedPropertyId();
            String matterIdValue = getPropertyString("matterIdValue");
            String principalUser = getPropertyString("principalUserEmail");

            if (extPropId == null || StringUtils.isBlank(matterIdValue) ||
                    StringUtils.isBlank(principalUser)) return 0;

            // Extract date filters only (subject ignored)
            String dateFrom = null;
            String dateTo   = null;
            String subject  = null;

            if (filters != null) {
                for (DataListFilterQueryObject f : filters) {
                    String q = (f.getQuery() != null) ? f.getQuery().toLowerCase() : "";
                    String[] vals = f.getValues();
                    if (vals == null || vals.length == 0) continue;
                    String v0 = vals[0];
                    if (StringUtils.isBlank(v0)) continue;

                    if (q.contains("subject")) {
                        subject = v0;
                    } else if (q.contains("start_filter") || q.contains("start")) {
                        dateFrom = toUtcIso(v0, false);
                    } else if (q.contains("end_filter") || q.contains("end")) {
                        dateTo = toUtcIso(v0, true);
                    }
                }
            }

            String url = buildCountUrl(principalUser, extPropId, matterIdValue,
                    dateFrom, dateTo,subject);
//            LogUtil.info(TAG, "[getCount] URL: " + url);

            String raw = executeGet(url, token);
            JsonNode root = MAPPER.readTree(raw);
            if (root.has("@odata.count")) {
                return root.get("@odata.count").asInt(0);
            }
            return 0;
        } catch (Exception e) {
            LogUtil.error(TAG, e, "[getCount] Error");
            return 0;
        }
    }

    // ----------------------------------------------------------------
    // URL builders (with optional date filters)
    // ----------------------------------------------------------------
    private String buildPrimaryUrl(String user, String extPropId,
                                   String matterIdValue,
                                   String dateFrom, String dateTo,
                                   String subject,
                                   int top, int skip) throws Exception {

        String escapedPropId = extPropId.replace("'", "''");
        String escapedMatter = matterIdValue.replace("'", "''");

        StringBuilder filter = new StringBuilder();
        filter.append("isDraft eq false and ");
        filter.append("singleValueExtendedProperties/any(ep: ep/id eq '")
                .append(escapedPropId).append("' and ep/value eq '")
                .append(escapedMatter).append("')");

        if (StringUtils.isNotBlank(dateFrom)) {
            filter.append(" and sentDateTime ge ").append(dateFrom);
        }
        if (StringUtils.isNotBlank(dateTo)) {
            filter.append(" and sentDateTime le ").append(dateTo);
        }



        if (StringUtils.isNotBlank(subject)) {
            String safeSubject = subject.replace("%", "").trim().replace("'", "''");
            if (StringUtils.isNotBlank(safeSubject)) {
                if (StringUtils.isBlank(dateFrom)
                        && StringUtils.isBlank(dateTo)) {
                    filter.append(" and sentDateTime ge 1900-01-01T00:00:00Z ");
                }
                filter.append(" and contains(subject,'")
                        .append(safeSubject)
                        .append("')");
            }
        }

        List<String> params = new ArrayList<>();
        params.add("$filter=" + enc(filter.toString()));
        params.add("$orderby=" + enc("sentDateTime desc"));
        params.add("$select=" + enc(
                "id,subject,from,toRecipients,sentDateTime,hasAttachments," +
                        "conversationId,internetMessageId"
        ));

        String expandFilter = "singleValueExtendedProperties($filter=id eq '" +
                escapedPropId + "')";
        params.add("$expand=" + enc(expandFilter));

        if (top  > 0) params.add("$top="  + top);
        if (skip > 0) params.add("$skip=" + skip);

        return GRAPH_BASE + enc(user) + "/mailFolders/SentItems/messages?" +
                String.join("&", params);
    }

    private String buildCountUrl(String user, String extPropId,
                                 String matterIdValue,
                                 String dateFrom, String dateTo, String subject) throws Exception {

        String escapedPropId = extPropId.replace("'", "''");
        String escapedMatter = matterIdValue.replace("'", "''");

        StringBuilder filter = new StringBuilder();
        filter.append("isDraft eq false and ");
        filter.append("singleValueExtendedProperties/any(ep: ep/id eq '")
                .append(escapedPropId).append("' and ep/value eq '")
                .append(escapedMatter).append("')");

        if (StringUtils.isNotBlank(dateFrom)) {
            filter.append(" and sentDateTime ge ").append(dateFrom);
        }
        if (StringUtils.isNotBlank(dateTo)) {
            filter.append(" and sentDateTime le ").append(dateTo);
        }

        if (StringUtils.isNotBlank(subject)) {
            String safeSubject = subject.replace("%", "").trim().replace("'", "''");
            if (StringUtils.isNotBlank(safeSubject)) {
                filter.append(" and contains(subject,'")
                        .append(safeSubject)
                        .append("')");
            }
        }



        List<String> params = new ArrayList<>();
        params.add("$filter=" + enc(filter.toString()));
        params.add("$count=true");
        params.add("$top=1");
        params.add("$select=id");

        return GRAPH_BASE + enc(user) + "/mailFolders/SentItems/messages?" +
                String.join("&", params);
    }

    // ----------------------------------------------------------------
    // Thread unread count
    // ----------------------------------------------------------------
    private int fetchUnreadCount(String user, String conversationId,
                                 String token) {
        try {
            String filter = "conversationId eq '" +
                    conversationId.replace("'", "''") + "' " +
                    "and isRead eq false and isDraft eq false";

            String url = GRAPH_BASE + enc(user) + "/messages" +
                    "?$filter=" + enc(filter) +
                    "&$count=true" +
                    "&$top=1" +
                    "&$select=id";

            String raw = executeGet(url, token);
            JsonNode root = MAPPER.readTree(raw);
            if (root.has("@odata.count")) {
                return root.get("@odata.count").asInt(0);
            }
            JsonNode value = root.get("value");
            return (value != null && value.isArray()) ? value.size() : 0;
        } catch (Exception e) {
            LogUtil.warn(TAG, "[fetchUnreadCount] conv=" + conversationId +
                    " err=" + e.getMessage());
            return 0;
        }
    }

    // ----------------------------------------------------------------
    // HTTP helpers
    // ----------------------------------------------------------------
    private String executeGet(String url, String token) throws Exception {
        return Request.Get(url)
                .addHeader("Authorization", "Bearer " + token)
                .addHeader("Accept", "application/json")
                .addHeader("ConsistencyLevel", "eventual")
                .execute()
                .returnContent()
                .asString();
    }

    private JsonNode fetchJsonArray(String url, String token) throws Exception {
        String raw = executeGet(url, token);
        JsonNode root = MAPPER.readTree(raw);
        if (root.has("error")) {
            LogUtil.error(TAG, null, "Graph API error: " + root.get("error"));
            return null;
        }
        return root.get("value");
    }

    // ----------------------------------------------------------------
    // Row builder
    // ----------------------------------------------------------------
    private FormRow toFormRow(JsonNode msg, int unreadCount) {
        FormRow row = new FormRow();
        row.setProperty("id",               text(msg, "id"));
        row.setProperty("subject",          text(msg, "subject"));
        row.setProperty("from",             extractFrom(msg));
        row.setProperty("toRecipients",     extractRecipients(msg.get("toRecipients")));
        row.setProperty("hasAttachments",   msg.path("hasAttachments").asBoolean(false) ? "Yes" : "No");
        row.setProperty("conversationId",   text(msg, "conversationId"));
        row.setProperty("internetMessageId",text(msg, "internetMessageId"));
        row.setProperty("matterId",         extractMatterId(msg));
        row.setProperty("unreadCount",      String.valueOf(unreadCount));
        row.setProperty("dateCreated",      formatDateTime(text(msg, "sentDateTime")));
        row.setProperty("internetMessageIdEncoded", text(msg, "internetMessageId").replace("<", "&lt;").replace(">", "&gt;"));
        return row;
    }

    // ----------------------------------------------------------------
    // Field extractors
    // ----------------------------------------------------------------
    private String extractFrom(JsonNode msg) {
        JsonNode from = msg.path("from").path("emailAddress");
        if (from.isMissingNode()) return "";
        String name = from.path("name").asText("");
        String addr = from.path("address").asText("");
        if (StringUtils.isNotBlank(name)) {
            return name + " <" + addr + ">";
        }
        return addr;
    }

    private String extractRecipients(JsonNode recipients) {
        if (recipients == null || !recipients.isArray()) return "";
        List<String> emails = new ArrayList<>();
        for (JsonNode rec : recipients) {
            String addr = rec.path("emailAddress").path("address").asText("");
            if (StringUtils.isNotBlank(addr)) emails.add(addr);
        }
        return String.join("; ", emails);
    }

    private String extractMatterId(JsonNode msg) {
        JsonNode props = msg.get("singleValueExtendedProperties");
        if (props != null && props.isArray()) {
            for (JsonNode p : props) {
                String val = p.path("value").asText("");
                if (StringUtils.isNotBlank(val)) return val;
            }
        }
        return "";
    }

    // ----------------------------------------------------------------
    // Utilities
    // ----------------------------------------------------------------
    private String text(JsonNode node, String field) {
        if (node == null) return "";
        JsonNode n = node.get(field);
        return (n == null || n.isNull()) ? "" : n.asText("");
    }

    private String formatDateTime(String raw) {
        if (StringUtils.isBlank(raw)) return "";
        try {
            ZonedDateTime zdt = ZonedDateTime.parse(raw,
                    DateTimeFormatter.ISO_DATE_TIME);
            return zdt.format(DISPLAY_FMT);
        } catch (Exception e) {
            return raw;
        }
    }

    private String enc(String s) throws Exception {
        return URLEncoder.encode(s, StandardCharsets.UTF_8.name());
    }

    private String toUtcIso(String raw, boolean isEndOfDay) {
        if (StringUtils.isBlank(raw)) return null;
        raw = raw.trim();
        if (raw.length() > 10 && raw.contains("T") && (raw.endsWith("Z") || raw.contains("+"))) {
            return raw;
        }
        if (raw.contains(" ")) raw = raw.substring(0, raw.indexOf(' '));
        if (raw.contains("T")) raw = raw.substring(0, raw.indexOf('T'));
        return isEndOfDay ? raw + "T23:59:59Z" : raw + "T00:00:00Z";
    }

    private String buildExtendedPropertyId() {
        String guid = StringUtils.trimToNull(getPropertyString("extendedPropertyGuid"));
        String name = StringUtils.trimToNull(getPropertyString("extendedPropertyName"));
        if (guid == null || name == null) return null;
        return "String {" + guid + "} Name " + name;
    }

    // ----------------------------------------------------------------
    // App‑only Access Token
    // ----------------------------------------------------------------
    private String getAppOnlyAccessToken() throws Exception {
        String clientId     = StringUtils.trimToNull(getPropertyString("clientId"));
        String clientSecret = StringUtils.trimToNull(getPropertyString("clientSecret"));
        String tenantId     = StringUtils.trimToNull(getPropertyString("tenantId"));

        if (clientId == null || clientSecret == null || tenantId == null) {
            LogUtil.warn(TAG, "Missing OAuth credentials");
            return null;
        }

        String tokenUrl = TOKEN_BASE + tenantId + "/oauth2/v2.0/token";

        Content response = Request.Post(tokenUrl)
                .bodyForm(Form.form()
                        .add("client_id",     clientId)
                        .add("client_secret", clientSecret)
                        .add("scope",         GRAPH_SCOPE)
                        .add("grant_type",    "client_credentials")
                        .build())
                .execute()
                .returnContent();

        JsonNode json = MAPPER.readTree(response.asString());
        String token = json.path("access_token").asText("");
        if (StringUtils.isBlank(token)) {
            LogUtil.error(TAG, null, "Access token missing in response");
            return null;
        }
        return token;
    }
}