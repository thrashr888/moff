use serde_json::{Value, json};

use crate::operations::{delete_message_request, field, single_item};
use crate::{
    OutlookWebError, OutlookWebSession, ProductWrite, find_item_request, get_item_request, header,
};

fn contact(value: &Value) -> Result<&Value, OutlookWebError> {
    if value["__type"] != "Contact:#Exchange" {
        return Err(OutlookWebError::UnsupportedItem);
    }
    Ok(value)
}

fn normalized(item: &Value) -> Result<Value, OutlookWebError> {
    contact(item)?;
    let mut value = json!({"id":field(item,"/ItemId/Id")?});
    for (graph, exchange) in text_fields() {
        value[graph] = item.get(exchange).cloned().unwrap_or_else(|| json!(""));
    }
    value["emailAddresses"] = json!([]);
    if let Some(emails) = item.get("EmailAddresses").and_then(Value::as_array) {
        value["emailAddresses"] = json!(
            ["EmailAddress1", "EmailAddress2", "EmailAddress3"]
                .iter()
                .filter_map(|key| emails.iter().find(|email| email["Key"] == *key))
                .map(|email| json!({"name":email["Name"],"address":email["EmailAddress"]}))
                .collect::<Vec<_>>()
        );
    }
    for (graph, keys) in [
        ("businessPhones", vec!["BusinessPhone", "BusinessPhone2"]),
        ("homePhones", vec!["HomePhone", "HomePhone2"]),
    ] {
        let values =
            item.get("PhoneNumbers")
                .and_then(Value::as_array)
                .map_or(Vec::new(), |phones| {
                    keys.iter()
                        .filter_map(|key| {
                            phones
                                .iter()
                                .find(|phone| phone["Key"] == *key)
                                .and_then(|phone| phone["PhoneNumber"].as_str())
                        })
                        .collect::<Vec<_>>()
                });
        value[graph] = json!(values);
    }
    value["mobilePhone"] = json!(
        item.get("PhoneNumbers")
            .and_then(Value::as_array)
            .and_then(|phones| phones.iter().find(|phone| phone["Key"] == "MobilePhone"))
            .and_then(|phone| phone["PhoneNumber"].as_str())
            .unwrap_or("")
    );
    Ok(value)
}

fn text_fields() -> [(&'static str, &'static str); 8] {
    [
        ("givenName", "GivenName"),
        ("surname", "Surname"),
        ("displayName", "DisplayName"),
        ("companyName", "CompanyName"),
        ("jobTitle", "JobTitle"),
        ("department", "Department"),
        ("officeLocation", "OfficeLocation"),
        ("personalNotes", "PersonalNotes"),
    ]
}

pub fn validate_contact_input(input: &Value) -> Result<(), OutlookWebError> {
    fields(input).map(|_| ())
}

fn verify(input: &Value, after: &Value) -> bool {
    input.as_object().is_some_and(|fields| {
        fields.iter().all(|(key, expected)| {
            if key == "emailAddresses" {
                let Some(expected) = expected.as_array() else {
                    return false;
                };
                let Some(actual) = after[key].as_array() else {
                    return false;
                };
                expected.len() == actual.len()
                    && expected.iter().zip(actual).all(|(expected, actual)| {
                        expected["address"] == actual["address"]
                            && expected
                                .get("name")
                                .is_none_or(|name| name == &actual["name"])
                    })
            } else {
                expected == &after[key]
            }
        })
    })
}

fn fields(input: &Value) -> Result<Value, OutlookWebError> {
    ProductWrite::contact(None, "POST", Some(input.clone()), None)?;
    let mut item = json!({"__type":"Contact:#Exchange"});
    for (graph, exchange) in text_fields() {
        if let Some(value) = input.get(graph) {
            item[exchange] = value.clone();
        }
    }
    if let Some(emails) = input.get("emailAddresses") {
        let emails = emails
            .as_array()
            .ok_or(OutlookWebError::InvalidInput("invalid email addresses"))?;
        if emails.len() > 3 {
            return Err(OutlookWebError::InvalidInput(
                "Outlook contacts support at most three email addresses",
            ));
        }
        item["EmailAddresses"] = json!(emails.iter().enumerate().map(|(i,email)|
            json!({"Key":format!("EmailAddress{}",i+1),"EmailAddress":email["address"],"Name":email.get("name").and_then(Value::as_str).unwrap_or(""),"RoutingType":"SMTP"})).collect::<Vec<_>>());
    }
    let mut phones = Vec::new();
    for (name, prefix) in [
        ("businessPhones", "BusinessPhone"),
        ("homePhones", "HomePhone"),
    ] {
        if let Some(values) = input.get(name) {
            let values = values
                .as_array()
                .ok_or(OutlookWebError::InvalidInput("invalid phone list"))?;
            if values.len() > 2 {
                return Err(OutlookWebError::InvalidInput(
                    "Outlook contacts support at most two business/home phone numbers",
                ));
            }
            for (index, value) in values.iter().enumerate() {
                phones.push(json!({"Key":if index == 0 {prefix.to_owned()} else {format!("{prefix}2")},"PhoneNumber":value}));
            }
        }
    }
    if let Some(value) = input.get("mobilePhone") {
        phones.push(json!({"Key":"MobilePhone","PhoneNumber":value}));
    }
    if !phones.is_empty()
        || input.get("businessPhones").is_some()
        || input.get("homePhones").is_some()
    {
        item["PhoneNumbers"] = json!(phones);
    }
    Ok(item)
}

impl OutlookWebSession {
    pub fn contacts_list(&self, limit: u16, offset: u32) -> Result<Value, OutlookWebError> {
        if !(1..=100).contains(&limit) {
            return Err(OutlookWebError::InvalidInput("contact limit must be 1-100"));
        }
        let mut request = find_item_request(limit);
        request["Body"]["ParentFolderIds"][0]["Id"] = json!("contacts");
        request["Body"]["Paging"]["Offset"] = json!(offset);
        request["Body"]["ItemShape"]["BaseShape"] = json!("AllProperties");
        request["Body"]["ItemShape"]["BodyType"] = json!("Text");
        let response = self.call("FindItem", &request)?;
        let root = response
            .pointer("/Body/ResponseMessages/Items/0/RootFolder")
            .ok_or(OutlookWebError::InvalidResponse)?;
        let items = root["Items"]
            .as_array()
            .ok_or(OutlookWebError::InvalidResponse)?;
        let items = items
            .iter()
            .map(normalized)
            .collect::<Result<Vec<_>, _>>()?;
        let last = root["IncludesLastItemInRange"]
            .as_bool()
            .ok_or(OutlookWebError::InvalidResponse)?;
        let next = root["IndexedPagingOffset"]
            .as_u64()
            .ok_or(OutlookWebError::InvalidResponse)?;
        Ok(json!({"value":items,"count":items.len(),"has_more":!last,"next_offset":next}))
    }

    pub fn contact_get(&self, id: &str) -> Result<Value, OutlookWebError> {
        if id.trim().is_empty() {
            return Err(OutlookWebError::InvalidInput(
                "contact ID must not be empty",
            ));
        }
        normalized(single_item(&self.call("GetItem", &get_item_request(id))?)?)
    }

    pub fn contacts_search(&self, query: &str, limit: u16) -> Result<Value, OutlookWebError> {
        if query.trim().is_empty() || !(1..=100).contains(&limit) {
            return Err(OutlookWebError::InvalidInput(
                "contacts search requires a nonempty query and limit 1-100",
            ));
        }
        let query = query.to_lowercase();
        let mut offset = 0;
        let mut items = Vec::new();
        for _ in 0..100 {
            let page = self.contacts_list(100, offset)?;
            let values = page["value"]
                .as_array()
                .ok_or(OutlookWebError::InvalidResponse)?;
            items.extend(
                values
                    .iter()
                    .filter(|item| item.to_string().to_lowercase().contains(&query))
                    .cloned(),
            );
            if items.len() >= usize::from(limit) || page["has_more"] == false {
                let more = items.len() > usize::from(limit) || page["has_more"] == true;
                items.truncate(usize::from(limit));
                return Ok(
                    json!({"value":items,"count":items.len(),"has_more":more,"complete":!more}),
                );
            }
            let next = page["next_offset"]
                .as_u64()
                .and_then(|n| u32::try_from(n).ok())
                .ok_or(OutlookWebError::InvalidResponse)?;
            if next <= offset {
                return Err(OutlookWebError::InvalidResponse);
            }
            offset = next;
        }
        Err(OutlookWebError::TruncatedSearch)
    }

    pub fn contact_create(&self, input: &Value) -> Result<Value, OutlookWebError> {
        let item = fields(input)?;
        let request = json!({"__type":"CreateItemJsonRequest:#Exchange","Header":header(),"Body":{
            "__type":"CreateItemRequest:#Exchange","MessageDisposition":"SaveOnly",
            "SavedItemFolderId":{"__type":"TargetFolderId:#Exchange","BaseFolderId":{"__type":"DistinguishedFolderId:#Exchange","Id":"contacts"}},
            "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},"Items":[item]}});
        let response = self.execute("CreateItem", &request, true)?;
        let id = field(single_item(&response)?, "/ItemId/Id")?;
        let mut after = self
            .contact_get(id)
            .map_err(|_| OutlookWebError::ItemVerificationFailed { id: id.to_owned() })?;
        if !verify(input, &after) {
            return Err(OutlookWebError::ItemVerificationFailed { id: id.to_owned() });
        }
        after["verified"] = json!(true);
        Ok(after)
    }

    pub fn contact_update(&self, id: &str, input: &Value) -> Result<Value, OutlookWebError> {
        let fields = fields(input)?;
        let before = self.call("GetItem", &get_item_request(id))?;
        let before = contact(single_item(&before)?)?;
        let key = field(before, "/ItemId/ChangeKey")?;
        let mut updates = Vec::new();
        for (name, value) in fields.as_object().ok_or(OutlookWebError::InvalidResponse)? {
            if name == "__type" {
                continue;
            }
            if matches!(name.as_str(), "EmailAddresses" | "PhoneNumbers") {
                let (dictionary, keys) = if name == "EmailAddresses" {
                    (
                        "EmailAddress",
                        vec!["EmailAddress1", "EmailAddress2", "EmailAddress3"],
                    )
                } else {
                    (
                        "PhoneNumber",
                        vec![
                            "BusinessPhone",
                            "BusinessPhone2",
                            "HomePhone",
                            "HomePhone2",
                            "MobilePhone",
                        ],
                    )
                };
                for index in keys {
                    if name == "PhoneNumbers"
                        && !((index.starts_with("Business")
                            && input.get("businessPhones").is_some())
                            || (index.starts_with("Home") && input.get("homePhones").is_some())
                            || (index == "MobilePhone" && input.get("mobilePhone").is_some()))
                    {
                        continue;
                    }
                    let path = json!({"__type":"DictionaryPropertyUri:#Exchange","FieldURI":dictionary,"FieldIndex":index});
                    let entry = value
                        .as_array()
                        .ok_or(OutlookWebError::InvalidResponse)?
                        .iter()
                        .find(|entry| entry["Key"] == index);
                    if let Some(entry) = entry {
                        let mut item = json!({"__type":"Contact:#Exchange"});
                        item[name] = json!([entry]);
                        updates.push(
                            json!({"__type":"SetItemField:#Exchange","Path":path,"Item":item}),
                        );
                    } else {
                        updates.push(json!({"__type":"DeleteItemField:#Exchange","Path":path}));
                    }
                }
                continue;
            }
            let mut item = json!({"__type":"Contact:#Exchange"});
            item[name] = value.clone();
            updates.push(json!({"__type":"SetItemField:#Exchange","Path":{"__type":"PropertyUri:#Exchange","FieldURI":name},"Item":item}));
        }
        let request = json!({"__type":"UpdateItemJsonRequest:#Exchange","Header":header(),"Body":{
            "__type":"UpdateItemRequest:#Exchange","MessageDisposition":"SaveOnly","ConflictResolution":"NeverOverwrite","SuppressReadReceipts":true,
            "ItemShape":{"__type":"ItemResponseShape:#Exchange","BaseShape":"IdOnly"},
            "ItemChanges":[{"__type":"ItemChange:#Exchange","ItemId":{"__type":"ItemId:#Exchange","Id":id,"ChangeKey":key},"Updates":updates}]}});
        self.execute("UpdateItem", &request, true)?;
        let mut after = self
            .contact_get(id)
            .map_err(|_| OutlookWebError::ItemVerificationFailed { id: id.to_owned() })?;
        if !verify(input, &after) {
            return Err(OutlookWebError::ItemVerificationFailed { id: id.to_owned() });
        }
        after["verified"] = json!(true);
        Ok(after)
    }

    pub fn contact_delete(&self, id: &str) -> Result<Value, OutlookWebError> {
        let before = self.call("GetItem", &get_item_request(id))?;
        let item = contact(single_item(&before)?)?;
        self.execute(
            "DeleteItem",
            &delete_message_request(id, field(item, "/ItemId/ChangeKey")?)?,
            true,
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::test_support::{message, mock};

    #[test]
    fn verification_handles_cleared_fields_and_native_dictionary_order() {
        let after = normalized(&json!({"__type":"Contact:#Exchange","ItemId":{"Id":"id"},
            "EmailAddresses":[{"Key":"EmailAddress2","Name":"Second","EmailAddress":"b@example.invalid"},
                {"Key":"EmailAddress1","Name":"First","EmailAddress":"a@example.invalid"}]})).unwrap();
        assert!(verify(
            &json!({"givenName":"","homePhones":[],"mobilePhone":"",
            "emailAddresses":[{"address":"a@example.invalid"},{"address":"b@example.invalid","name":"Second"}]}),
            &after
        ));
        assert!(!verify(&json!({"givenName":"Wrong"}), &after));
        assert!(!verify(&json!({"emailAddresses":[]}), &after));
    }

    #[test]
    fn contact_inputs_reject_unknown_fields_and_slot_overflows() {
        for input in [
            json!({"owners":[]}),
            json!({"businessPhones":["1","2","3"]}),
            json!({"emailAddresses":[{"address":"a@example.com"},{"address":"b@example.com"},{"address":"c@example.com"},{"address":"d@example.com"}]}),
        ] {
            assert!(validate_contact_input(&input).is_err());
        }
        let item = fields(&json!({"givenName":"Test","emailAddresses":[{"address":"test@example.com"}],"mobilePhone":"123"})).unwrap();
        assert_eq!(item["__type"], "Contact:#Exchange");
        assert_eq!(item["EmailAddresses"][0]["Key"], "EmailAddress1");
        assert_eq!(item["PhoneNumbers"][0]["Key"], "MobilePhone");
    }

    #[test]
    fn indexed_contact_updates_preserve_unrelated_phone_categories_and_verify() {
        let before = json!({"__type":"Contact:#Exchange","ItemId":{"Id":"id","ChangeKey":"key"},
            "GivenName":"Test","PhoneNumbers":[{"Key":"HomePhone","PhoneNumber":"home"},{"Key":"BusinessPhone","PhoneNumber":"old"}]});
        let after = json!({"__type":"Contact:#Exchange","ItemId":{"Id":"id","ChangeKey":"new"},
            "GivenName":"Test","PhoneNumbers":[{"Key":"HomePhone","PhoneNumber":"home"},{"Key":"BusinessPhone","PhoneNumber":"new"}]});
        let (session, requests, completion) = mock(vec![
            ("GetItem", message(before)),
            ("UpdateItem", message(json!({"ItemId":{"Id":"id"}}))),
            ("GetItem", message(after)),
        ]);
        let result = session
            .contact_update("id", &json!({"businessPhones":["new"]}))
            .unwrap();
        assert_eq!(result["verified"], true);
        assert_eq!(result["homePhones"], json!(["home"]));
        completion.join().unwrap();
        let requests = requests.lock().unwrap();
        let changes = requests[1]["Body"]["ItemChanges"][0]["Updates"]
            .as_array()
            .unwrap();
        assert_eq!(changes.len(), 2);
        assert_eq!(
            changes[0]["Path"]["__type"],
            "DictionaryPropertyUri:#Exchange"
        );
        assert_eq!(changes[0]["Path"]["FieldIndex"], "BusinessPhone");
        assert_eq!(changes[1]["__type"], "DeleteItemField:#Exchange");
        assert_eq!(
            requests[1]["Body"]["ItemChanges"][0]["ItemId"]["ChangeKey"],
            "key"
        );
        assert_eq!(requests[1]["Body"]["ConflictResolution"], "NeverOverwrite");
    }

    #[test]
    fn created_contacts_are_saved_only_to_contacts_and_failed_verification_retains_id() {
        let (session, requests, completion) = mock(vec![
            ("CreateItem", message(json!({"ItemId":{"Id":"created"}}))),
            (
                "GetItem",
                message(
                    json!({"__type":"Contact:#Exchange","ItemId":{"Id":"created"},"GivenName":"wrong"}),
                ),
            ),
        ]);
        let error = session
            .contact_create(&json!({"givenName":"expected"}))
            .unwrap_err();
        assert!(error.to_string().contains("created"));
        completion.join().unwrap();
        let requests = requests.lock().unwrap();
        assert_eq!(
            requests[0]["Body"]["SavedItemFolderId"]["BaseFolderId"]["Id"],
            "contacts"
        );
        assert_eq!(requests[0]["Body"]["MessageDisposition"], "SaveOnly");
    }

    #[test]
    fn contact_search_follows_sparse_offsets_and_rejects_noncontact_deletion() {
        let root = |items: Value, last: bool, offset: u32| {
            json!({"Body":{"ResponseMessages":{"Items":[{
            "ResponseCode":"NoError","RootFolder":{"Items":items,"IncludesLastItemInRange":last,"IndexedPagingOffset":offset}}]}}})
        };
        let (session, _, completion) = mock(vec![
            ("FindItem", root(json!([]), false, 1)),
            (
                "FindItem",
                root(
                    json!([{"__type":"Contact:#Exchange","ItemId":{"Id":"id"},"GivenName":"Target"}]),
                    true,
                    2,
                ),
            ),
        ]);
        assert_eq!(session.contacts_search("target", 1).unwrap()["count"], 1);
        completion.join().unwrap();
        let (session, _, completion) = mock(vec![(
            "GetItem",
            message(json!({"__type":"Message:#Exchange","ItemId":{"Id":"id","ChangeKey":"key"}})),
        )]);
        assert!(session.contact_delete("id").is_err());
        completion.join().unwrap();
    }
}
