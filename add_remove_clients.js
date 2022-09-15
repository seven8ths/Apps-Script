function addGroupMember(tableId, rowId, groupKey, client) {
  // Get the table row data
  const rowName = "tables/" + tableId + "/rows/" + rowId;
  const row = Area120Tables.Tables.Rows.get(rowName);

  const member = {
    email: client,
    role: "MEMBER",
    delivery_settings: "NONE",
  };

  try {
    AdminDirectory.Members.insert(member, groupKey);
    Logger.log("User %s added as a member of group %s.", member, groupKey);
    row.values["Group"] = groupKey;
    Area120Tables.Tables.Rows.patch(row, rowName);
  } catch (err) {
    Logger.log("Failed with error %s", err.message);
  }
}

function removeGroupMember(tableId, rowId, groupKey, client) {
  // Get the table row data
  const rowName = "tables/" + tableId + "/rows/" + rowId;
  const row = Area120Tables.Tables.Rows.get(rowName);

  try {
    AdminDirectory.Members.remove(groupKey, client);
    Logger.log("User %s removed as a member of group %s.", client, groupKey);
    row.values["Group"] = "";
    Area120Tables.Tables.Rows.patch(row, rowName);
  } catch (err) {
    Logger.log("Failed with error %s", err.message);
  }
}
