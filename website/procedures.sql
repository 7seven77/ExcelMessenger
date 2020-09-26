PROCEDURE `AddMessage`(IN `senderIn` TEXT, IN `recipientIn` TEXT, IN `messageIn` TEXT)
BEGIN

-- The number of messages two people can have stored on the database
-- Change this value to whatever you want
SELECT 15 INTO @maxMessages;

-- Get the number of messages between the two people
SELECT COUNT(*) INTO @numberOfMessages FROM `MESSAGE`
WHERE (`sender` = senderIn and `recipient` = recipientIn) or  (`sender` = recipientIn and `recipient` = senderIn);

-- If the number of messages is more than the number of messages allowed,
-- find the first message that was sent (Least recent) and get it's date
-- If the limit is not exceeded, it is set to NULL
IF @numberOfMessages >= @maxMessages THEN
	SELECT MIN(`date`) INTO @smallestDate FROM `MESSAGE`
	WHERE (`sender` = senderIn and `recipient` = recipientIn) or  (`sender` = recipientIn and `recipient` = senderIn);
ELSE
	SELECT NULL INTO @smallestDate;
END IF;

-- Remove the least recent message before adding the new one
DELETE FROM `MESSAGE` WHERE `date` = @smallestDate;

INSERT INTO `MESSAGE`(`sender`, `recipient`, `message`, `date`)
VALUES (senderIn, recipientIn, messageIn, NOW());

END