/**
 * Subroute Label Provider.
 * 
 * (C) Bartosz Sławecki (@bswck)
 */

// IMPORTANT: Security settings.
const LABEL_CREATORS = [
  "bswck.dev",
]

// Handy constants.
const EMAIL_AT = "@";
const EMAIL_TAG = "+";
const LABEL_SEP = "+";
const LABEL_SUBLABEL = ".";
const LABEL_SPACE = "_";
const RECIPIENT_PATTERN = '^".+" <(.+)>$';

// Utilities.
const identity = obj => obj;
const makeTitle = string => string[0].toUpperCase() + string.substring(1).toLowerCase();
const repr = JSON.stringify;

/**
 * Parse an e-mail address into two parts: localPart and domain.
 *
 * @function getEmailParts
 * @private
 * @param  {String} emailAddress
 *     A valid e-mail address.
 * @return {localPart: String, domain: String}    
 *     A dictionary with e-mail parts.
 */
function getEmailParts(emailAddress) {
  const parts = emailAddress.split(EMAIL_AT);
  const domain = parts.pop();
  const localPart = parts.join(EMAIL_AT);
  return { localPart, domain };
}

/**
 * Parse the subroute of an e-mail address.
 * 
 * For instance, the subroute of
 * "jonathan+joestar@unlikeable.animes.com"
 * is "joestar".
 *
 * @function getSubroute
 * @private
 * @param  {String} emailAddress
 *     A valid e-mail address.
 * @return {String}
 *     The subroute.
 */
function getSubroute(emailAddress) {
  const localPart = getEmailParts(emailAddress).localPart;
  const tagAt = localPart.indexOf(EMAIL_TAG);
  return (tagAt == -1) ? "direct" : localPart.slice(tagAt + 1);
}

/**
 * Parse the labels indicated by a subroute of an e-mail.
 * 
 * For instance, the labels of a subroute 'abc.def+foo_bar+biz'
 * are 'Abc/Def', 'Foo Bar' and 'Biz'.
 *
 * @function getLabelsFromSubroute
 * @private
 * @param  {String} subroute
 *     A subroute to parse labels of.
 * @return {Array<GmailApp.GmailThread>}
 *     The labels from the subroute.
 */
function getLabelsFromSubroute(subroute, createIfNecessary = false) {
  const labelNames = subroute.split(LABEL_SEP).map(
    subroutePart => subroutePart.split(LABEL_SUBLABEL).map(
      tag => tag.split(LABEL_SPACE).map(makeTitle).join(" "),
    ).join("/")
  );
  return labelNames.map(
    labelName => getLabel(labelName, createIfNecessary),
  ).filter(identity);
}

/**
 * Gets a label from the GmailApp service.
 * 
 * If it doesn't exist *and `createIfNecessary` is true*,
 * creates and returns a new label with the specified name.
 *
 * @function getLabel
 * @private
 * @param  {Boolean} createIfNecessary
 *     Whether to create the label if it doesn't exist.
 * @return {GmailApp.GmailLabel?}
 *     The fetched label.
 */
function getLabel(labelName, createIfNecessary = false) {
  return (
    GmailApp.getUserLabelByName(labelName)
    || createIfNecessary && GmailApp.createLabel(labelName)
  );
}

/**
 * Extracts the e-mail from e-mail recipient signatures.
 * For instance, the e-mail extracted from
 *   the signature '"Display Name" <actual@email.com>'
 *   is 'actual@email.com'.
 *
 * @function extractAddress
 * @private
 * @param  {String} emailRecipient
 *     Any string, possibly matching
 *     the aforementioned signature format.
 * @return {String}
 *     A clean e-mail address from the signature.
 */
function extractAddress(emailRecipient) {
  const matches = emailRecipient.match(RECIPIENT_PATTERN);
  return matches ? matches[1] : emailRecipient;
}

/**
 * Adds subroute-determined labels to a thread.
 *
 * @function addLabelsToThread
 * @param  {GmailApp.GmailThread} thread
 *     A Gmail Thread to add labels to.
 */
function addLabelsToThread(thread) {
  for (
    const [senderAddress, recipientAddress]
    of thread.getMessages().map(
      message => [
        message.getFrom(),
        message.getTo()
      ].map(extractAddress)
    )
  ) {
    if (!(senderAddress && recipientAddress)) continue;

    const subroute = getSubroute(recipientAddress);
    const labels = getLabelsFromSubroute(
      subroute,
      LABEL_CREATORS.includes(getEmailParts(senderAddress).domain),
    ).filter(label => !thread.getLabels().includes(label));

    if (labels.length) {
      const labelNames = labels.map(
        label => label.getName(),
      ).map(repr).join(", ");
      labels.map(label => thread.addLabel(label));

      console.log(
        "Added label(s) %s to thread %s from",
        labelNames,
        repr(thread.getFirstMessageSubject()),
        thread.getMessages()[0].getFrom(),
      );
    }
    break;
  }
}

/**
 * Serves as the entrypoint for the Subroute Label Provider add-on.
 * Walks through the inbox threads and adds labels
 *   based on the subroutes of the recipient addresses.
 * For instance, an e-mail to casino.omega+inquiries@gmail.com
 * would get an "Inquires" label in the casino.omega@gmail.com's inbox.
 *
 * @function addLabelsToInboxMessages
 */
function addLabelsToInboxMessages() {
  GmailApp.getInboxThreads().map(addLabelsToThread);
}
