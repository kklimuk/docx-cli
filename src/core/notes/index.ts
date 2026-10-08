export { type NoteConfig, type NoteKind, noteConfig } from "./config";
export {
	NoteBody,
	NoteReferenceRun,
	TrackedNoteBody,
	wrapNoteBodyAsDeleted,
	wrapNoteBodyAsEdited,
} from "./emit";
export { buildEmptyNotesTree } from "./empty";
export {
	insertNoteReferenceAtOffset,
	NoteOffsetOutOfRangeError,
	paragraphNoteLength,
	removeNoteReferences,
	wrapNoteReferencesAsDeleted,
} from "./splice";
