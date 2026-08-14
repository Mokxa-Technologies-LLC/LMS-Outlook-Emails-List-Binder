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
 * OutlookSentListBinder
 *
 * Lists sent emails for a given user.
 * Supports server-side date range filtering via hidden DataList filter columns:
 *   - start_filter  → sentDateTime ge
 *   - end_filter    → sentDateTime le
 * Subject filtering is applied client-side after fetching.
 * Uses app-only (client credentials) access token.
 */
public class OutlookSentListBinder extends DataListBinderDefault {

    private static final String TAG         = "OutlookSentListBinder";
    private static final String GRAPH_BASE  = "https://graph.microsoft.com/v1.0/users/";
    private static final String TOKEN_BASE  = "https://login.microsoftonline.com/";
    private static final String GRAPH_SCOPE = "https://graph.microsoft.com/.default";
    private static final ObjectMapper MAPPER = new ObjectMapper();
    private static final DateTimeFormatter DISPLAY_FMT =
            DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    // Plugin metadata
    @Override public String getName()        { return getClass().getName(); }
    @Override public String getVersion()     { return "1.0.0"; }
    @Override public String getLabel()       { return "Outlook Sent List (Date+Subject Filter)"; }
    @Override public String getDescription() {
        return "Lists sent emails for a user. " +
                "Supports server-side date filters (start_filter / end_filter). " +
                "Subject filtering is applied client-side.";
    }

    @Override
    public String getPropertyOptions() {
        return AppUtil.readPluginResource(getClassName(),
                "/properties/OutlookSentListBinder.json", null, true, null);
    }

    // Columns
    @Override
    public DataListColumn[] getColumns() {
        return new DataListColumn[]{
                new DataListColumn("id",                "Message ID",          false),
                new DataListColumn("subject",           "Subject",             true),
                new DataListColumn("from",              "From",                true),
                new DataListColumn("toRecipients",      "To",                  true),
                new DataListColumn("sentDate",          "Sent Date",           true),
                new DataListColumn("hasAttachments",    "Attachments",         true),
                new DataListColumn("importance",        "Importance",          true),
                new DataListColumn("unreadCount",       "Unread",              true),
                new DataListColumn("conversationId",    "Conversation ID",     false),
                new DataListColumn("internetMessageId", "Internet Message ID", false),
                new DataListColumn("bodyPreview",       "Preview",             true),
                new DataListColumn("start_filter",      "Start Filter",        true),
                new DataListColumn("end_filter",        "End Filter",          true),
                new DataListColumn("extendedProperty",  "Extended Property",   true),
                new DataListColumn("internetMessageIdEncoded", "Internet Message ID Encoded", false),

        };
    }

    @Override public String getPrimaryKeyColumnName() { return "id"; }

    // getData
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

            String principalUser = getPropertyString("principalUserEmail");
            if (StringUtils.isBlank(principalUser)) {
                LogUtil.warn(TAG, "Missing principalUserEmail property.");
                return result;
            }

            // Extract filters
            String dateFrom = null;
            String dateTo   = null;
            String subject  = null;

            if (filters != null) {
                for (DataListFilterQueryObject f : filters) {
                    String q      = (f.getQuery() != null) ? f.getQuery().toLowerCase() : "";
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

            String url = buildSentUrl(principalUser, dateFrom, dateTo, pageSize, skip);
//            LogUtil.info(TAG, "[getData] URL: " + url);

            JsonNode messages = fetchJsonArray(url, token);
            if (messages == null) return result;

            // Client-side subject filter
            List<JsonNode> msgList = new ArrayList<>();
            for (JsonNode msg : messages) {
                msgList.add(msg);
            }

            if (StringUtils.isNotBlank(subject)) {
                String term = subject.toLowerCase().replace("%", "").trim();
                List<JsonNode> filtered = new ArrayList<>();
                for (JsonNode msg : msgList) {
                    String subj = text(msg, "subject");
                    if (subj != null && subj.toLowerCase().contains(term)) {
                        filtered.add(msg);
                    }
                }
                msgList = filtered;
            }

            FormRowSet rowSet = new FormRowSet();
            for (JsonNode msg : msgList) {
                String convId = text(msg, "conversationId");

                int[]    unreadOut  = {0};
                String[] extPropOut = {""};

                if (StringUtils.isNotBlank(convId)) {
                    //fetchConversationData(principalUser, convId, token, buildExtendedPropId(), unreadOut, extPropOut);
                }
                rowSet.add(toFormRow(msg, unreadOut[0], extPropOut[0]));
            }
//            LogUtil.info(TAG, "[getData] RowSet: " + rowSet);
            result.addAll(rowSet);

        } catch (Exception e) {
            LogUtil.error(TAG, e, "[getData] Error");
        }
        return result;
    }

    // getDataTotalRowCount
    @Override
    public int getDataTotalRowCount(DataList dataList, Map properties,
                                    DataListFilterQueryObject[] filters) {
        try {
            String token = getAppOnlyAccessToken();
            if (token == null) return 0;

            String principalUser = getPropertyString("principalUserEmail");
            if (StringUtils.isBlank(principalUser)) return 0;

            String dateFrom = null;
            String dateTo   = null;

            if (filters != null) {
                for (DataListFilterQueryObject f : filters) {
                    String q      = (f.getQuery() != null) ? f.getQuery().toLowerCase() : "";
                    String[] vals = f.getValues();
                    if (vals == null || vals.length == 0) continue;
                    String v0 = vals[0];
                    if (StringUtils.isBlank(v0)) continue;

                    if (q.contains("start_filter")) {
                        dateFrom = toUtcIso(v0, false);
                    } else if (q.contains("end_filter")) {
                        dateTo   = toUtcIso(v0, true);
                    }
                }
            }

            String url = buildCountUrl(principalUser, dateFrom, dateTo);
//            LogUtil.info(TAG, "[getCount] URL: " + url);

            String   raw  = executeGet(url, token);
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

    // URL builders
    private String buildSentUrl(String user, String dateFrom, String dateTo,
                                int top, int skip) throws Exception {

        StringBuilder filter = new StringBuilder();

        if (StringUtils.isNotBlank(dateFrom)) {
            filter.append("sentDateTime ge ").append(dateFrom);
        }
        if (StringUtils.isNotBlank(dateTo)) {
            if (filter.length() > 0) filter.append(" and ");
            filter.append("sentDateTime le ").append(dateTo);
        }

        List<String> params = new ArrayList<>();
        if (filter.length() > 0) {
            params.add("$filter=" + enc(filter.toString()));
        }
        params.add("$orderby=" + enc("sentDateTime desc"));
        params.add("$select=" + enc(
                "id,subject,from,toRecipients,sentDateTime," +
                        "hasAttachments,importance,conversationId," +
                        "internetMessageId,bodyPreview"
        ));
        if (top  > 0) params.add("$top="  + top);
        if (skip > 0) params.add("$skip=" + skip);

        return GRAPH_BASE + enc(user) + "/mailFolders/SentItems/messages?" +
                String.join("&", params);
    }

    private String buildCountUrl(String user, String dateFrom,
                                 String dateTo) throws Exception {

        StringBuilder filter = new StringBuilder();

        if (StringUtils.isNotBlank(dateFrom)) {
            filter.append("sentDateTime ge ").append(dateFrom);
        }
        if (StringUtils.isNotBlank(dateTo)) {
            if (filter.length() > 0) filter.append(" and ");
            filter.append("sentDateTime le ").append(dateTo);
        }

        List<String> params = new ArrayList<>();
        if (filter.length() > 0) {
            params.add("$filter=" + enc(filter.toString()));
        }
        params.add("$count=true");
        params.add("$top=1");
        params.add("$select=id");

        return GRAPH_BASE + enc(user) + "/mailFolders/SentItems/messages?" +
                String.join("&", params);
    }

    // HTTP helpers
    private String executeGet(String url, String token) throws Exception {
        return Request.Get(url)
                .addHeader("Authorization", "Bearer " + token)
                .addHeader("Accept",        "application/json")
                .addHeader("ConsistencyLevel", "eventual")
                .execute()
                .returnContent()
                .asString();
    }

    private JsonNode fetchJsonArray(String url, String token) throws Exception {
        String   raw  = executeGet(url, token);
        JsonNode root = MAPPER.readTree(raw);
        if (root.has("error")) {
            LogUtil.error(TAG, null, "Graph API error: " + root.get("error"));
            return null;
        }
        return root.get("value");
    }

    // Row builder
    private FormRow toFormRow(JsonNode msg, int unreadCount, String extendedProperty) {
        //LogUtil.info(TAG, msg.toString());
        FormRow row = new FormRow();
        row.setProperty("id",                text(msg, "id"));
        row.setProperty("subject",           text(msg, "subject"));
        row.setProperty("from",              extractFrom(msg));
        row.setProperty("toRecipients",      extractRecipients(msg.get("toRecipients")));
        row.setProperty("sentDate",          formatDateTime(text(msg, "sentDateTime")));
        row.setProperty("hasAttachments",    msg.path("hasAttachments").asBoolean(false) ? "Yes" : "No");
        row.setProperty("importance",        capitalize(text(msg, "importance")));
        row.setProperty("conversationId",    text(msg, "conversationId"));
        row.setProperty("internetMessageId", text(msg, "internetMessageId"));
        row.setProperty("bodyPreview",       text(msg, "bodyPreview"));
        row.setProperty("unreadCount",       String.valueOf(unreadCount));
        row.setProperty("extendedProperty",  extendedProperty);
        row.setProperty("internetMessageIdEncoded", text(msg, "internetMessageId").replace("<", "&lt;").replace(">", "&gt;"));
        //LogUtil.info(TAG, text(msg, "internetMessageId"));


        return row;
    }

    // Field extractors
    private String extractFrom(JsonNode msg) {
        JsonNode from = msg.path("from").path("emailAddress");
        if (from.isMissingNode()) return "";
        String name = from.path("name").asText("");
        String addr = from.path("address").asText("");
        return StringUtils.isNotBlank(name) ? name + " <" + addr + ">" : addr;
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

    // Utilities
    private String text(JsonNode node, String field) {
        if (node == null) return "";
        JsonNode n = node.get(field);
        return (n == null || n.isNull()) ? "" : n.asText("");
    }

    private String formatDateTime(String raw) {
        if (StringUtils.isBlank(raw)) return "";
        try {
            return ZonedDateTime.parse(raw, DateTimeFormatter.ISO_DATE_TIME)
                    .format(DISPLAY_FMT);
        } catch (Exception e) {
            return raw;
        }
    }

    private String capitalize(String s) {
        if (StringUtils.isBlank(s)) return "";
        return Character.toUpperCase(s.charAt(0)) + s.substring(1).toLowerCase();
    }

    private String enc(String s) throws Exception {
        return URLEncoder.encode(s, StandardCharsets.UTF_8.name());
    }

    private String toUtcIso(String raw, boolean isEndOfDay) {
        if (StringUtils.isBlank(raw)) return null;
        raw = raw.trim();
        if (raw.length() > 10 && raw.contains("T") &&
                (raw.endsWith("Z") || raw.contains("+"))) {
            return raw;
        }
        if (raw.contains(" ")) raw = raw.substring(0, raw.indexOf(' '));
        if (raw.contains("T")) raw = raw.substring(0, raw.indexOf('T'));
        return isEndOfDay ? raw + "T23:59:59Z" : raw + "T00:00:00Z";
    }

    // App-only Access Token (client credentials)
    private String getAppOnlyAccessToken() throws Exception {
        String clientId     = StringUtils.trimToNull(getPropertyString("clientId"));
        String clientSecret = StringUtils.trimToNull(getPropertyString("clientSecret"));
        String tenantId     = StringUtils.trimToNull(getPropertyString("tenantId"));

        if (clientId == null || clientSecret == null || tenantId == null) {
            LogUtil.warn(TAG, "Missing OAuth credentials (clientId / clientSecret / tenantId)");
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

        JsonNode json  = MAPPER.readTree(response.asString());
        String   token = json.path("access_token").asText("");
        if (StringUtils.isBlank(token)) {
            LogUtil.error(TAG, null, "Access token missing in response");
            return null;
        }
        return token;
    }

    private int fetchUnreadCount(String user, String conversationId, String token) {
        try {
            String filter = "conversationId eq '" +
                    conversationId.replace("'", "''") + "' " +
                    "and isRead eq false and isDraft eq false";

            String url = GRAPH_BASE + enc(user) + "/messages" +
                    "?$filter=" + enc(filter) +
                    "&$count=true" +
                    "&$top=1" +
                    "&$select=id";

            String   raw  = executeGet(url, token);
            JsonNode root = MAPPER.readTree(raw);
            if (root.has("@odata.count")) {
                return root.get("@odata.count").asInt(0);
            }
            JsonNode value = root.get("value");
            return (value != null && value.isArray()) ? value.size() : 0;
        } catch (Exception e) {
            LogUtil.warn(TAG, "[fetchUnreadCount] conv=" + conversationId + " err=" + e.getMessage());
            return 0;
        }
    }


    private String getExtendedPropertyFromConversation(String user,
                                                       String conversationId,
                                                       String token,
                                                       String extPropId) {
        try {

            if (StringUtils.isBlank(conversationId)
                    || StringUtils.isBlank(extPropId)) {
                return "";
            }

            String expand =
                    "singleValueExtendedProperties($filter=id eq '" +
                            extPropId.replace("'", "''") + "')";

            String url = GRAPH_BASE + enc(user) + "/messages" +
                    "?$filter=" + enc(
                    "conversationId eq '" +
                            conversationId.replace("'", "''") + "'"
            ) +
                    "&$select=" + enc("id,singleValueExtendedProperties") +
                    "&$expand=" + enc(expand);

//            LogUtil.info(TAG,
//                    "[getExtendedPropertyFromConversation] url=" + url);

            String raw = executeGet(url, token);

            JsonNode root = MAPPER.readTree(raw);

            if (root.has("error")) {
                LogUtil.warn(TAG,
                        "[getExtendedPropertyFromConversation] Graph Error: "
                                + root.get("error"));
                return "";
            }

            JsonNode messages = root.path("value");

            if (!messages.isArray()) {
                return "";
            }

            for (JsonNode msg : messages) {

                JsonNode props =
                        msg.path("singleValueExtendedProperties");

                if (!props.isArray()) {
                    continue;
                }

                for (JsonNode prop : props) {

                    String value =
                            prop.path("value").asText("").trim();

                    if (StringUtils.isNotBlank(value)) {

//                        LogUtil.info(TAG,
//                                "[getExtendedPropertyFromConversation] Found value="
//                                        + value
//                                        + ", conversationId="
//                                        + conversationId);

                        return value;
                    }
                }
            }

        } catch (Exception e) {
            LogUtil.warn(TAG,
                    "[getExtendedPropertyFromConversation] conversationId="
                            + conversationId
                            + ", error="
                            + e.getMessage());
        }

        return "";
    }



    private void fetchConversationData(String user, String conversationId,
                                       String token, String extPropId,
                                       int[] unreadOut, String[] extPropOut) {
        unreadOut[0]  = 0;
        extPropOut[0] = "";
        try {
            if (StringUtils.isBlank(conversationId)) return;

            String expand = "singleValueExtendedProperties($filter=id eq '" +
                    extPropId.replace("'", "''") + "')";

            String url = GRAPH_BASE + enc(user) + "/messages" +
                    "?$filter=" + enc("conversationId eq '" +
                    conversationId.replace("'", "''") + "'") +
                    "&$select=" + enc("id,isRead,isDraft,singleValueExtendedProperties") +
                    "&$expand=" + enc(expand) +
                    "&$top=50";

            String   raw   = executeGet(url, token);
            JsonNode root  = MAPPER.readTree(raw);
            if (root.has("error")) {
                LogUtil.warn(TAG, "[fetchConversationData] error=" + root.get("error"));
                return;
            }

            JsonNode messages = root.path("value");
            if (!messages.isArray()) return;

            for (JsonNode msg : messages) {
                // count unread (exclude drafts)
                boolean isRead  = msg.path("isRead").asBoolean(true);
                boolean isDraft = msg.path("isDraft").asBoolean(false);
                if (!isRead && !isDraft) {
                    unreadOut[0]++;
                }

                // grab extended property if not found yet
                if (extPropOut[0].isEmpty()) {
                    JsonNode props = msg.path("singleValueExtendedProperties");
                    if (props.isArray()) {
                        for (JsonNode prop : props) {
                            String val = prop.path("value").asText("").trim();
                            if (StringUtils.isNotBlank(val)) {
                                extPropOut[0] = val;
                                break;
                            }
                        }
                    }
                }
            }

        } catch (Exception e) {
            LogUtil.warn(TAG, "[fetchConversationData] conv=" + conversationId +
                    " err=" + e.getMessage());
        }
    }


    private String buildExtendedPropId() {

        String guid = getPropertyString("extendedPropertyGuid");
        String name = getPropertyString("extendedPropertyName");

        if (guid == null || guid.isEmpty()) {
            return null;
        }

        if (name == null || name.isEmpty()) {
            return null;
        }

        return "String {" + guid + "} Name " + name;
    }
}