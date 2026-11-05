# End-of-action annotation site

A web page where annotators mark the frame at which an action ends, on videos stored in Google Drive. Results go to a Google Sheet and to `annotations.csv` in the same Drive folder as `extensions.csv`.

Annotators see only the action name and a running number. The page never receives a clip name or which version it's showing.

| File | What it is |
|---|---|
| `index.html`, `config.js` | The page, hosted on GitHub Pages |
| `apps_script/Code.gs` | The backend, which runs in the results Sheet |
| `tests/code_gs.test.js` | Backend tests: `node tests/code_gs.test.js` |

**This repository holds code only.** Never commit videos, CSV files, the Drive folder id or the API key.

## Set it up

You need a Drive folder containing `extensions.csv` and the action folders (`<Action>/original <Action>/` and `<Action>/synthesized <Action>/`), shared as "Anyone with the link: Viewer".

### 1. Create an API key
1. In [Google Cloud Console](https://console.cloud.google.com), create a project.
2. Under **APIs & Services > Library**, enable **Google Drive API**.
3. Under **APIs & Services > Credentials**, choose **Create credentials > API key**.
4. Edit the key. Under **Application restrictions**, choose **Websites** and add `https://<your-github-name>.github.io/*`. Under **API restrictions**, choose **Google Drive API**.

### 2. Set up the backend
1. Create a Google Sheet for the results. Keep it private, because it holds the clip key.
2. In the Sheet, open **Extensions > Apps Script** and replace the contents with `apps_script/Code.gs`.
3. At the top of the script, fill in `ROOT_FOLDER_ID`, `DRIVE_API_KEY` and `ACCESS_CODE`, a code you choose for all annotators. These stay in Apps Script and don't go back into this repository.
4. Select `setup` in the toolbar and click **Run**. Approve the permissions; Google warns that the app is unverified, so choose **Advanced**, then continue. If any videos are missing from Drive, the error lists them and nothing changes.

### 3. Deploy
1. In Apps Script, choose **Deploy > New deployment > Web app**. Set **Execute as** to **Me** and **Who has access** to **Anyone**.
2. Copy the web app URL into `config.js`, then commit and push.
3. In the GitHub repository, go to **Settings > Pages** and deploy from the `main` branch, `/ (root)` folder.

The site appears at `https://<your-github-name>.github.io/<repository>/`.

### 4. Invite annotators
Send each person the site address and the access code. They type their own name on the start page, and a new name starts a new session.

## Results

The `annotations` tab has one row per annotator and clip, and `annotations.csv` in Drive is rewritten after every saved mark. Edit neither by hand, because the next save rebuilds the CSV from the tab, and the page shows annotators the marks held in the `sessions` tab.

| Column | Meaning |
|---|---|
| `annotator`, `action`, `clip_id` | Who marked which clip, and the action it's filed under |
| `original_end_s`, `synthesized_end_s`, `extension_s` | Copied from `extensions.csv` |
| `annotation_original_s`, `annotation_original_frame` | The mark on the original |
| `annotation_synthesized_s`, `annotation_synthesized_frame` | The mark on the synthesized version |
| `annotated_action_original`, `annotated_action_synthesized` | The action name the annotator gave each version |
| `flag_original`, `flag_synthesized` | `no_clear_end` or `video_problem` if chosen |
| `annotated_at` | Time of the latest mark for that clip, in UTC |

Frames count from 0. The page shows frames counting from 1, so frame 12 on screen is saved as frame 11.

## Making changes

- **After replacing a video in Drive:** run `setup` again. It rebuilds the clip key and keeps all sessions and results.
- **After changing `Code.gs`:** use **Deploy > Manage deployments**, edit the deployment and choose **New version**. Otherwise the old code keeps running.

## Limits

- **Names:** the name is how an annotator resumes, ignoring capitals. Two people who type the same name share one session and overwrite each other's marks, and a different spelling starts a new session, so ask for full names.
- **Blinding:** an annotator who opens the browser's developer tools can see a video's Drive file id, and opening that file in Drive shows its name. Don't give annotators the Drive folder link.
- **Who can get the videos:** the Drive API key only reads files shared with "Anyone with the link", so anyone with that link or a file id can also download them.
- **Download size:** each video downloads in full before it plays, about 39 MB on average, or about 4.3 GB per annotator for the 110-video set.
- **Google's limits:** Drive can temporarily refuse files that are downloaded heavily. The page then shows **Try loading again**.
