import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import logger from './logger.js';
import GraphClient from './graph-client.js';

interface ContentItem {
  type: 'text';
  text: string;
  [key: string]: unknown;
}

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;
  [key: string]: unknown;
}

/**
 * Register composite tools that chain multiple Graph API calls
 * These provide higher-level abstractions for common workflows
 */
export function registerCompositeTools(
  server: McpServer,
  graphClient: GraphClient
): void {
  logger.info('Registering composite tools...');

  // get-transcript-by-meeting: Fetch meeting transcript in one call
  server.tool(
    'get-transcript-by-meeting',
    `Retrieve a Teams meeting transcript by searching for the meeting by date and subject.

This composite tool chains multiple Graph API calls internally:
1. Queries calendar for meetings on the specified date
2. Finds the meeting matching the subject
3. Resolves the online meeting ID via join URL
4. Fetches the transcript content

Returns the VTT transcript with speaker attribution and timestamps.

💡 TIP: Works regardless of whether you organized the meeting or were just an attendee.`,
    {
      date: z
        .string()
        .describe('Date to search for meetings (ISO format: YYYY-MM-DD, e.g., "2026-01-16")'),
      subjectContains: z
        .string()
        .describe('Partial text to match in meeting subject (case-insensitive)'),
      startTime: z
        .string()
        .describe('Optional: Start time to narrow search (HH:MM format, e.g., "12:00")')
        .optional(),
      endTime: z
        .string()
        .describe('Optional: End time to narrow search (HH:MM format, e.g., "13:00")')
        .optional(),
      timezone: z
        .string()
        .describe('IANA timezone (e.g., "America/Chicago"). Defaults to UTC.')
        .optional(),
    },
    {
      title: 'get-transcript-by-meeting',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async (params) => getTranscriptByMeeting(graphClient, params)
  );

  logger.info('Composite tools registered successfully');
}

interface GetTranscriptParams {
  date: string;
  subjectContains: string;
  startTime?: string;
  endTime?: string;
  timezone?: string;
}

async function getTranscriptByMeeting(
  graphClient: GraphClient,
  params: GetTranscriptParams
): Promise<CallToolResult> {
  const { date, subjectContains, startTime, endTime, timezone = 'UTC' } = params;

  try {
    logger.info(`get-transcript-by-meeting: Searching for "${subjectContains}" on ${date}`);

    // Step 1: Build date range for calendar query
    const startDateTime = startTime
      ? `${date}T${startTime}:00`
      : `${date}T00:00:00`;
    const endDateTime = endTime
      ? `${date}T${endTime}:00`
      : `${date}T23:59:59`;

    // Step 2: Query calendar view
    // Use makeRequest directly to get raw JSON (graphRequest formats to TOON)
    logger.info(`Step 1: Querying calendar from ${startDateTime} to ${endDateTime}`);
    const calendarPath = `/me/calendarView?startDateTime=${encodeURIComponent(startDateTime)}&endDateTime=${encodeURIComponent(endDateTime)}`;
    const calendarOptions: Record<string, unknown> = {
      headers: {
        Prefer: `outlook.timezone="${timezone}"`,
      },
    };

    let calendarData: any;
    try {
      calendarData = await graphClient.makeRequest(calendarPath, calendarOptions);
    } catch (error) {
      return {
        content: [{ type: 'text', text: `Failed to query calendar: ${(error as Error).message}` }],
        isError: true,
      };
    }

    const events = calendarData.value || [];

    if (events.length === 0) {
      return {
        content: [{ type: 'text', text: `No meetings found on ${date}` }],
        isError: true,
      };
    }

    // Step 3: Find matching meeting by subject
    const matchingEvent = events.find((event: any) =>
      event.subject?.toLowerCase().includes(subjectContains.toLowerCase()) &&
      event.isOnlineMeeting === true
    );

    if (!matchingEvent) {
      const availableMeetings = events
        .filter((e: any) => e.isOnlineMeeting)
        .map((e: any) => `- ${e.subject} (${e.start?.dateTime})`)
        .join('\n');

      return {
        content: [{
          type: 'text',
          text: `No online meeting found matching "${subjectContains}".\n\nAvailable online meetings on ${date}:\n${availableMeetings || '(none)'}`
        }],
        isError: true,
      };
    }

    logger.info(`Step 2: Found meeting "${matchingEvent.subject}"`);

    // Step 4: Get full event details to extract joinUrl
    const eventPath = `/me/events/${matchingEvent.id}?$select=subject,onlineMeeting,onlineMeetingUrl,isOnlineMeeting`;
    let eventData: any;
    try {
      eventData = await graphClient.makeRequest(eventPath, {});
    } catch (error) {
      return {
        content: [{ type: 'text', text: `Failed to get event details: ${(error as Error).message}` }],
        isError: true,
      };
    }

    const joinUrl = eventData.onlineMeeting?.joinUrl;

    if (!joinUrl) {
      return {
        content: [{ type: 'text', text: `Meeting "${matchingEvent.subject}" does not have a Teams join URL` }],
        isError: true,
      };
    }

    logger.info(`Step 3: Got join URL, resolving meeting ID`);

    // Step 5: Query online meetings by join URL to get meeting ID
    const meetingsPath = `/me/onlineMeetings?$filter=JoinWebUrl eq '${encodeURIComponent(joinUrl)}'`;
    let meetingsData: any;
    try {
      meetingsData = await graphClient.makeRequest(meetingsPath, {});
    } catch (error) {
      return {
        content: [{ type: 'text', text: `Failed to resolve meeting ID: ${(error as Error).message}` }],
        isError: true,
      };
    }

    const onlineMeeting = meetingsData.value?.[0];

    if (!onlineMeeting) {
      return {
        content: [{ type: 'text', text: `Could not resolve online meeting for "${matchingEvent.subject}"` }],
        isError: true,
      };
    }

    const meetingId = onlineMeeting.id;
    logger.info(`Step 4: Resolved meeting ID, fetching transcripts`);

    // Step 6: List transcripts for this meeting
    const transcriptsPath = `/me/onlineMeetings/${meetingId}/transcripts`;
    let transcriptsData: any;
    try {
      transcriptsData = await graphClient.makeRequest(transcriptsPath, {});
    } catch (error) {
      return {
        content: [{ type: 'text', text: `Failed to list transcripts: ${(error as Error).message}` }],
        isError: true,
      };
    }

    const transcripts = transcriptsData.value || [];

    if (transcripts.length === 0) {
      return {
        content: [{
          type: 'text',
          text: `No transcripts found for meeting "${matchingEvent.subject}".\n\nNote: Transcription must be enabled during the meeting to generate transcripts.`
        }],
        isError: true,
      };
    }

    // Use the first (usually only) transcript
    const transcript = transcripts[0];
    const transcriptId = transcript.id;
    logger.info(`Step 5: Found transcript, fetching content`);

    // Step 7: Fetch transcript content
    // makeRequest returns { message: 'OK!', rawResponse: text } for non-JSON responses
    const contentPath = `/me/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`;
    let contentData: any;
    try {
      contentData = await graphClient.makeRequest(contentPath, {
        headers: {
          Accept: 'text/vtt',
        },
      });
    } catch (error) {
      return {
        content: [{ type: 'text', text: `Failed to fetch transcript content: ${(error as Error).message}` }],
        isError: true,
      };
    }

    logger.info(`Step 6: Successfully retrieved transcript`);

    // makeRequest wraps non-JSON responses in { message: 'OK!', rawResponse: text }
    const vttContent = contentData.rawResponse || contentData.message || 'No content';

    // Build result with metadata
    const result = {
      meeting: {
        subject: matchingEvent.subject,
        start: matchingEvent.start,
        end: matchingEvent.end,
        organizer: onlineMeeting.participants?.organizer?.upn || 'unknown',
      },
      transcript: {
        id: transcriptId,
        createdDateTime: transcript.createdDateTime,
        endDateTime: transcript.endDateTime,
      },
      content: vttContent,
    };

    return {
      content: [{
        type: 'text',
        text: JSON.stringify(result, null, 2),
      }],
    };

  } catch (error) {
    logger.error(`get-transcript-by-meeting error: ${(error as Error).message}`);
    return {
      content: [{
        type: 'text',
        text: `Error retrieving transcript: ${(error as Error).message}`,
      }],
      isError: true,
    };
  }
}
