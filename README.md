# End-of-action annotation tool

A web tool for marking the frame where an action ends in first-person maintenance videos.

Annotators enter their name and an access code, then step through the videos one at a time and mark the first frame where the named action is physically finished. They can move frame by frame, correct the action name, or flag a video that has no clear end or has a problem. Progress saves after every video, so they can stop and resume later.

Each clip appears twice, as the original and as a version extended past the end of the action. The two versions are shuffled and kept far apart, and annotators aren't told which one they're watching. This makes it possible to compare the marks on both versions.

Videos are stored in Google Drive, and results are saved to a Google Sheet and a CSV file in Drive.
