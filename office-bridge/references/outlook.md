# Outlook Patterns

Outlook uses `Office.context.mailbox` directly - no `.run()` pattern like other apps.

## Getting Sessions

```typescript
const bridge = await connect();
const sessions = await bridge.outlook();  // Returns OutlookSession[]
const mail = sessions[0];
```

## Execution Context

Code has access to `Office`, `mailbox`, and `item`:

```javascript
// mailbox = Office.context.mailbox
// item = Office.context.mailbox.item (current message/appointment)
return item.subject;
```

## Read vs Compose Mode

Outlook behaves differently depending on mode:

- **Read mode**: Viewing a received message/appointment
- **Compose mode**: Writing a new message/appointment

Properties are accessed differently in each mode.

## Common Patterns

### Get Subject (Read Mode)

```javascript
return item.subject;  // Direct property access
```

### Get Subject (Compose Mode)

```javascript
return new Promise((resolve) => {
  item.subject.getAsync((result) => {
    resolve(result.value);
  });
});
```

### Get Sender Info (Read Mode)

```javascript
return {
  from: item.from?.emailAddress,
  displayName: item.from?.displayName,
  subject: item.subject
};
```

### Get Recipients (Compose Mode)

```javascript
return new Promise((resolve) => {
  item.to.getAsync((result) => {
    resolve(result.value.map(r => r.emailAddress));
  });
});
```

### Get Message Body (Read Mode)

```javascript
return new Promise((resolve) => {
  item.body.getAsync(Office.CoercionType.Text, (result) => {
    resolve(result.value);
  });
});
```

### Set Subject (Compose Mode)

```javascript
return new Promise((resolve) => {
  item.subject.setAsync("New Subject", (result) => {
    resolve(result.status === Office.AsyncResultStatus.Succeeded);
  });
});
```

### Set Body (Compose Mode)

```javascript
return new Promise((resolve) => {
  item.body.setAsync(
    "Hello,\n\nThis is the message body.",
    { coercionType: Office.CoercionType.Text },
    (result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded);
    }
  );
});
```

### Add Recipient (Compose Mode)

```javascript
return new Promise((resolve) => {
  item.to.addAsync(
    [{ emailAddress: "user@example.com", displayName: "User Name" }],
    (result) => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded);
    }
  );
});
```

### Get Attachments (Read Mode)

```javascript
return item.attachments.map(a => ({
  name: a.name,
  type: a.attachmentType,
  size: a.size
}));
```

### Get Item Type

```javascript
return {
  type: item.itemType,  // "message" or "appointment"
  mode: typeof item.subject === 'string' ? 'read' : 'compose'
};
```

### Get User's Email

```javascript
return mailbox.userProfile.emailAddress;
```

## Tips

- Read mode properties are synchronous, compose mode uses callbacks
- Use `Office.CoercionType.Text` or `Office.CoercionType.Html` for body
- `item` is only available when viewing/composing a message
- Appointment items have different properties (start, end, location, etc.)
- Wrap callbacks in Promises for cleaner async/await usage
