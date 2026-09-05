import { google } from "googleapis";
import {
  Application,
  Assignment,
  AssignmentStatus,
  Review,
  ReviewScore,
  OverallScore,
  Decision,
  DecisionValue,
} from "./types";

let cachedClient: ReturnType<typeof google.sheets> | null = null;

function getSheetsClient() {
  if (cachedClient) return cachedClient;

  const auth = new google.auth.JWT({
    email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  cachedClient = google.sheets({
    version: "v4",
    auth,
  });

  return cachedClient;
}

// READ FUNCTIONS

export async function getRows(
  spreadsheetId: string,
  range: string
): Promise<string[][]> {
  const sheets = getSheetsClient();

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range,
  });

  return res.data.values ?? [];
}

function rowsToObjects<T>(
  rows: string[][],
  mapRow: (get: (header: string) => string) => T,
  dataStartsAt: number = 1
): T[] {
  const keyRow = rows[0] ?? [];
  const headers = keyRow.map((h) =>
    h.trim().toLowerCase()
  );

  const dataRows = rows.slice(dataStartsAt);

  return dataRows.map((row) => {
    const get = (header: string) => {
      const index = headers.indexOf(
        header.trim().toLowerCase()
      );

      return index === -1 ? "" : row[index] ?? "";
    };

    return mapRow(get);
  });
}

export async function getApplications(): Promise<Application[]> {
  const rows = await getRows(
    process.env.SHEET_ID!,
    "Applications"
  );

  const applications =
    rowsToObjects<Application>(
      rows,
      (get) => ({
        timestamp: get("timestamp"),
        email: get("email"),
        name: get("name"),
        uscId: get("usc_id").trim(),
        phone: get("phone"),
        pronouns: get("pronouns"),
        year: get("year"),
        majorMinor: get("major_minor"),
        coursework: get("coursework"),
        howHeard: get("how_heard"),
        links: get("links"),
        whyJoin: get("why_join"),
        aiResponse: get("ai_response"),
        experienceResponse: get(
          "experience_response"
        ),
        socialResponse: get(
          "social_response"
        ),
        passionResponse: get(
          "passion_response"
        ),
        availability: get(
          "meeting_availability"
        ),
        maybeExplanation: get(
          "maybe_explanation"
        ),
        notifsReaction: get(
          "notifs_reaction"
        ),
      }),
      3
    );

  // Remove empty rows.
  const validApplications =
    applications.filter(
      (application) =>
        application.uscId.length > 0
    );

  // One USC ID = one application.
  //
  // Map.set() overwrites an earlier entry with a later one,
  // so if someone submits the form twice, we keep their
  // most recent submission.
  const applicationsByUscId =
    new Map<string, Application>();

  for (const application of validApplications) {
    applicationsByUscId.set(
      application.uscId,
      application
    );
  }

  return Array.from(
    applicationsByUscId.values()
  );
}

export async function getAssignments(): Promise<Assignment[]> {
  const rows = await getRows(
    process.env.SHEET_ID!,
    "Assignments"
  );

  return rowsToObjects<Assignment>(
    rows,
    (get) => ({
      uscId: get("usc_id"),
      reviewerEmail: get("reviewer_email"),
      status: get("status") as AssignmentStatus,
    })
  );
}

export async function getReviews(): Promise<Review[]> {
  const rows = await getRows(
    process.env.SHEET_ID!,
    "Reviews"
  );

  return rowsToObjects<Review>(
    rows,
    (get) => ({
      uscId: get("usc_id"),
      reviewerEmail: get("reviewer_email"),

      experienceScore: Number(
        get("experience_score")
      ) as ReviewScore,

      researchScore: Number(
        get("research_score")
      ) as ReviewScore,

      qualityScore: Number(
        get("quality_score")
      ) as ReviewScore,

      overallScore: Number(
        get("overall_score")
      ) as OverallScore,

      notes: get("notes"),

      submittedAt: get("submitted_at"),
    })
  );
}

// WRITE FUNCTIONS

export async function submitReview(
  review: Review
): Promise<void> {
  const sheets = getSheetsClient();

  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.SHEET_ID!,
    range: "Reviews",
    valueInputOption: "RAW",

    requestBody: {
      values: [
        [
          review.uscId,
          review.reviewerEmail,
          review.experienceScore,
          review.researchScore,
          review.qualityScore,
          review.overallScore,
          review.notes,
          review.submittedAt,
        ],
      ],
    },
  });
}

export async function updateReview(
  review: Review
): Promise<boolean> {
  const rows = await getRows(
    process.env.SHEET_ID!,
    "Reviews"
  );

  const headers = (rows[0] ?? []).map(
    (header) => header.trim().toLowerCase()
  );

  const uscIdCol =
    headers.indexOf("usc_id");

  const emailCol =
    headers.indexOf("reviewer_email");

  const dataIndex = rows
    .slice(1)
    .findIndex(
      (row) =>
        row[uscIdCol] === review.uscId &&
        row[emailCol] === review.reviewerEmail
    );

  if (dataIndex === -1) {
    return false;
  }

  const sheetRow = dataIndex + 2;

  const sheets = getSheetsClient();

  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SHEET_ID!,

    range: `Reviews!A${sheetRow}:H${sheetRow}`,

    valueInputOption: "RAW",

    requestBody: {
      values: [
        [
          review.uscId,
          review.reviewerEmail,
          review.experienceScore,
          review.researchScore,
          review.qualityScore,
          review.overallScore,
          review.notes,
          review.submittedAt,
        ],
      ],
    },
  });

  return true;
}

export async function createAssignment(
  uscId: string,
  reviewerEmail: string
): Promise<void> {
  const sheets = getSheetsClient();

  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.SHEET_ID!,
    range: "Assignments",
    valueInputOption: "RAW",
    requestBody: {
      values: [
        [
          uscId,
          reviewerEmail,
          "assigned",
        ],
      ],
    },
  });
}

export async function setAssignmentStatus(
  uscId: string,
  reviewerEmail: string,
  status: AssignmentStatus
): Promise<void> {
  const rows = await getRows(
    process.env.SHEET_ID!,
    "Assignments"
  );

  const headers = (rows[0] ?? []).map((h) =>
    h.trim().toLowerCase()
  );

  const uscIdCol = headers.indexOf("usc_id");
  const emailCol = headers.indexOf("reviewer_email");

  const dataIndex = rows
    .slice(1)
    .findIndex(
      (row) =>
        row[uscIdCol] === uscId &&
        row[emailCol] === reviewerEmail
    );

  if (dataIndex === -1) return;

  const sheetRow = dataIndex + 2;

  const sheets = getSheetsClient();

  await sheets.spreadsheets.values.update({
    spreadsheetId: process.env.SHEET_ID!,
    range: `Assignments!C${sheetRow}`,
    valueInputOption: "RAW",
    requestBody: {
      values: [[status]],
    },
  });
}

export async function getDecisions(): Promise<Decision[]> {
  const rows = await getRows(
    process.env.SHEET_ID!,
    "Decisions"
  );

  return rowsToObjects<Decision>(
    rows,
    (get) => ({
      uscId: get("usc_id"),
      decision: (
        get("decision") || "undecided"
      ) as DecisionValue,
      notes: get("notes"),
    })
  );
}

export async function setDecision(
  uscId: string,
  decision: DecisionValue,
  notes: string
): Promise<void> {
  const rows = await getRows(
    process.env.SHEET_ID!,
    "Decisions"
  );

  const headers = (rows[0] ?? []).map((h) =>
    h.trim().toLowerCase()
  );

  const uscIdCol = headers.indexOf("usc_id");

  const dataIndex = rows
    .slice(1)
    .findIndex(
      (row) => row[uscIdCol] === uscId
    );

  const sheets = getSheetsClient();

  if (dataIndex === -1) {
    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.SHEET_ID!,
      range: "Decisions",
      valueInputOption: "RAW",
      requestBody: {
        values: [[uscId, decision, notes]],
      },
    });
  } else {
    const sheetRow = dataIndex + 2;

    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.SHEET_ID!,
      range: `Decisions!A${sheetRow}:C${sheetRow}`,
      valueInputOption: "RAW",
      requestBody: {
        values: [[uscId, decision, notes]],
      },
    });
  }
}

// BULK ASSIGNMENT

export async function createAssignments(
  assignments: {
    uscId: string;
    reviewerEmail: string;
  }[]
): Promise<void> {
  if (assignments.length === 0) return;

  const sheets = getSheetsClient();

  await sheets.spreadsheets.values.append({
    spreadsheetId: process.env.SHEET_ID!,
    range: "Assignments",
    valueInputOption: "RAW",
    requestBody: {
      values: assignments.map(
        ({ uscId, reviewerEmail }) => [
          uscId,
          reviewerEmail,
          "assigned",
        ]
      ),
    },
  });
}