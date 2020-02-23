"use strict";

const needle = require("needle");
const Aria2 = require("aria2");
const fs = require("fs");
const path = require('path');


const user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.169 Safari/537.36";
//A browser user-agent, you can replace it with your own.
//https://www.whatismybrowser.com/  at the bottom is your user-agent.
//or get one from https://developers.whatismybrowser.com/useragents/explore/software_type_specific/web-browser/  (pick a very common one)

const aria2_connection_info_filepath = "aria2.conn.config";
//You need to put some configs in the filename/filepath specified above.

const tasklist_filepath = "tasklist";
//Where you put a list of sharepoint urls to download.
//Has to be written in specific format

const retry_minutes = 5;
//If you want more frequent trtries after error happens, tweak this value down.
//If you are downloading big sharepoint directory and you want the program to check progress less frequently, tweak this value up.
//5 minutes is not a definite standard to follow, I just happened to pick 5 minutes.

let aria2 = null;
let files_to_be_downloaded = [];
let GID_filepath_dictionary = {};
let aria2_max_concurrent_downloads = 1;
//This value will be overwritten by "max-concurrent-downloads" value in Aria2 server's config.

setTimeout(initialization, 0); //This is intended delay so "messages" would initialize
function initialization() {
	return new Promise((resolve, reject) => {
		//This part is the initialization process, should only run once.
		
		//load "downloader.config" file and check content.
		const aria2ConfigFileContent = fs.readFileSync(aria2_connection_info_filepath, "utf-8");
		let Aria2_IP = aria2ConfigFileContent.match(/Aria2_IP=(.*)/);
		let Aria2_port = aria2ConfigFileContent.match(/Aria2_port=(\d+)/);
		let Aria2_token = aria2ConfigFileContent.match(/Aria2_token=(.*)/);
		
		//checks for required config option.
		if (!Aria2_IP || Aria2_IP[1] == "")
			reject(messages.Aria2_IP_not_specified);
		else
			Aria2_IP = Aria2_IP[1];
		
		if (!Aria2_port || Aria2_IP[1] == 0)
			reject(messages.Aria2_port_not_specified);
		else
			Aria2_port = Aria2_port[1];
		
		if (!Aria2_token || Aria2_token[1] == "") {
			console.error(messages.Aria2_token_not_specified);
			Aria2_token = "";
		} else {
			Aria2_token = Aria2_token[1];
		}
		
		//create aria2 instance.
		console.log(messages.Aria2_connection_opening);
		aria2 = new Aria2({
			host: Aria2_IP,
			port: Aria2_port,
			secure: false,
			secret: Aria2_token,
			path: "/jsonrpc"
		});
		aria2.open().then(resolve);
	})
	.then(() => {
		//Register onDownloadStart event.
		//set "aria2_max_concurrent_downloads" to match Aria2 server's "max-concurrent-downloads" setting.
		console.log(messages.Aria2_connection_opened);
		return aria2.call("getGlobalOption")
			.then(options => {
				aria2_max_concurrent_downloads = options["max-concurrent-downloads"];
				console.log(messages.Aria2_concurrent_downloads(aria2_max_concurrent_downloads));
			})
	})
	.then(() => {
		//load linklist
		const linklist = fs.readFileSync(tasklist_filepath, "utf-8").replace(/\r/g, "").split("\n");
		if (linklist.length == 0)
			throw(messages.Invalid_linklist);
		
		const links_to_download = [];
		
		for (let i in linklist) {
			//Skip comments your might write in linklist file
			if (linklist[i].startsWith("#") || linklist[i].startsWith("//") || linklist[i] == "")
				continue;
			
			//The rest of for loop parses single line of linklist rule
			const regex_with_pwd = /^(https:\/\/[^\.]*\.sharepoint\.com[^ ]*) dir=="(.*)" pwd==(.*)$/;
			const regex_without_pwd = /^(https:\/\/[^\.]*\.sharepoint\.com[^ ]*) dir=="(.*)"$/;
			
			if (linklist[i].includes("onedrive.aspx")) {
				console.error(messages.invalid_linklist_syntax(i+1));
				continue;
			} else if (linklist[i].match(regex_without_pwd)) {
				//A linklist rule with password
				const tmp = linklist[i].match(regex_without_pwd);
				if (path.isAbsolute(tmp[2])) {
					tmp[2] = tmp[2].replace(/\\/g, "\/").replace(/\/$/,"");
					links_to_download.push(linkinfo(tmp[1], tmp[2]));
				} else {
					console.error(messages.not_absolute_path(tmp[2]));
				}
			} else if (linklist[i].match(regex_with_pwd)) {
				//A linklist rule without password
				const tmp = linklist[i].match(regex_with_pwd);
				if (path.isAbsolute(tmp[2])) {
					tmp[2] = tmp[2].replace(/\\/g, "\/").replace(/\/$/,"");
					links_to_download.push(linkinfo(tmp[1], tmp[2], tmp[3]));
				} else {
					console.error(messages.not_absolute_path(tmp[2]));
				}
			} else {
				console.error(messages.invalid_linklist_syntax(i+1));
				continue;
			}
		}
		if (links_to_download.length == 0)
			throw(messages.no_valid_sharepoint_link);
		console.log(messages.found_N_sharepoint_links(links_to_download.length));
		return(download_sharepoint_links(links_to_download));
	})
	.then(() => {
		console.log(messages.program_finished);
		//An attempt to grace disconnect from Aria2 server
		//Maybe it doesn't make a difference if I just calls process.exit directly
		//But I'm not sure.
		return aria2.close()
		.catch(()=>{})
		.then(() => {process.exit(0)});
	})
	.catch(error => {
		console.error(error);
		//An attempt to grace disconnect from Aria2 server
		//Maybe it doesn't make a difference if I just calls process.exit directly
		//But I'm not sure.
		if (aria2) {
			return aria2.close()
			.catch(()=>{})
			.then(() => {process.exit(1)});
		} else {
			process.exit(1);
		}
	})
}

function download_sharepoint_links(linkinfos) {
	//Iterate through every .sharepoint.com link and download them.
	let p = Promise.resolve();
	for (const info of linkinfos) {
		p = p.then(() => {
			return get_sharepoint_cookie(info);
		}).catch(() => {
			return;
		});
	}
	return p;
}

function get_sharepoint_cookie(linkinfo_) {
	return needle("get", linkinfo_.url)
	.then((response) => {
		//Handles redirection, set headers and cookies.
		let redirectURL = response.body.match(/Object moved to <a href="([^"]*)">/);
		if (response.body == "429 TOO MANY REQUESTS")
			throw(response.body);
		else if (response.body.includes('document.getElementById("txtPassword")')) {
			//This part handles password protected sharepoint folder.
			console.log(messages.folder_is_password_protected)
			if (linkinfo_.pwd) {
				const rawFormDataFields = response.body.match(/<input type="hidden" name="([^"]*)" id="\1" value="([^"]*)" \/>/g);
				const formDataFields = {};
				for (const field of rawFormDataFields) {
					const tmp = field.match(/<input type="hidden" name="([^"]*)" id="\1" value="([^"]*)" \/>/);
					formDataFields[tmp[1]] = tmp[2];
				}
				formDataFields["__EVENTTARGET"] = "btnSubmitPassword";
				formDataFields["__EVENTARGUMENT"] = "";
				formDataFields["txtPassword"] = linkinfo_.pwd;
				const posturl = linkinfo_.url.match(/(http.*\.sharepoint\.com)/)[1] + response.body.match(/action="([^"]*)"/)[1].replace("amp;","");
				
				return needle("post", posturl, formDataFields)
				.then(response => {
					if (response.body.includes('document.getElementById("txtPassword")')) {
						//User gave wrong password to program, so microsoft redirected us back to the password input page.
						throw (messages.password_incorrect(linkinfo_.url));
					} else {
						redirectURL = response.body.match(/Object moved to <a href="([^"]*)">/);
						const needle_options = {cookies: response.cookies, headers:{"user-agent": user_agent}};
						needle_options.cookies["FeatureOverrides_disableFeatures"] = "";
						needle_options.cookies["FeatureOverrides_enableFeatures"] = "";
						return download_single_sharepoint_link(redirectURL[1], linkinfo_.dir, needle_options);
					}
				})
			} else {
				throw(messages.password_not_provided(linkinfo_.url));
			}
		} else if (!redirectURL)
			throw(messages.redirect_url_not_found);
		else {
			const needle_options = {cookies: response.cookies, headers:{"user-agent": user_agent}};
			needle_options.cookies["FeatureOverrides_disableFeatures"] = "";
			needle_options.cookies["FeatureOverrides_enableFeatures"] = "";
			return download_single_sharepoint_link(redirectURL[1], linkinfo_.dir, needle_options);
		}
	})
	.catch((error) => {
		console.log(error);
		if (error != messages.sharepoint_link_download_finished &&
			error != messages.redirect_url_not_found &&
			!error.toString().startsWith(messages.password_incorrect_startswith) &&
			!error.toString().startsWith(messages.password_not_provided_startswith) &&
			!error.toString().startsWith(messages.unknown_status_code_startswith)) {
				console.log(messages.retry_in_N_minutes(retry_minutes));
				return timer()
					.then(() => get_sharepoint_cookie(linkinfo_))
		} else {
			return error;
		}
	});
}

function download_single_sharepoint_link(url, download_dir, needle_options) {
	return get_filelist_of_a_sharepoint_link(url, download_dir, needle_options)
	.catch(error => {
		throw(error);
	})
	.then(filelist => {
		console.log(messages.got_N_file_informations(filelist.length));
		files_to_be_downloaded = filelist;
		return loop_until_files_are_downloaded(needle_options, download_dir);
	})
}

function get_filelist_of_a_sharepoint_link(url, download_dir, needle_options, prefix="/") {
	if (prefix == "/")
		console.log(messages.getting_sharepoint_file_list);
	if (!url.includes("RowLimit="))
		url = url + "&RowLimit=9999"; //RowLimit parameter is added to get a list of more than 30 files;
	return needle("get", url, needle_options)
	.then(response => {
		const folderID = url.match("id=([^&]*)")[1];
		let urls = response.body.match(/"\.spItemUrl": "([^"]*)"/g).map(str =>
			sanitize_Unicode_encoded_string(str.match(/"\.spItemUrl": "([^"]*)"/)[1])
		);
		let filenames = response.body.match(/"FileLeafRef": "([^"]*)"/g).map(str =>
			sanitize_Unicode_encoded_string(str.match(/"FileLeafRef": "([^"]*)"/)[1])
		);
		let filesizes = response.body.match(/"FileSizeDisplay": "(.*)"/g).map(str => {
			if (str == '"FileSizeDisplay": ""')
				return 0;
			else
				return str.match(/"FileSizeDisplay": "(\d+)"/)[1]
		});
		//https://docs.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ee537053(v%3Doffice.14)
		let FSObjTypes = response.body.match(/"FSObjType": "([^"]*)"/g).map(str =>
			str.match(/"FSObjType": "(\d)"/)[1]
		);
		
		let files = [];
		let p = Promise.resolve();
		
		for (let i in FSObjTypes) {
			//p makes this for loop only continues after promises chain returned.
			//So that we are not trying to browse ALL the folders at the same time,
			//which will be slower, but will not make this program harmful to Microsoft servers.
			
			//This function could accidentally send 50 http requests at the same time.....before I tuned it,
			//One request at the time is too slow if you want to traverse large folder though...
			//At the moment I don't have a solution to traverse 2, 3 or more subfolders at the same time.
			if (FSObjTypes[i] == 1) {
				//console.log('"' + prefix + filenames[i] + '"' + " is a folder.");
				p = p.then(() => {return get_filelist_of_a_sharepoint_link(
					url.replace(folderID, folderID+encodeURIComponent("/"+filenames[i])),
					download_dir,
					needle_options,
					prefix + filenames[i] + "/"
				)})
				.then(filelist => {
					files = files.concat(filelist);
				})
				.catch(error => {
					throw(messages.folder_traverse_failed(prefix, error));
					//Any error happened will result in traversal of the whole sharepoint folder next try...
					//I have no idea on how to reserve already traversed folders for next time this function gets called...
				})
			} else if (FSObjTypes[i] == 0) {
				p = p.then(() => {
					files.push(fileInfo(download_dir + prefix + filenames[i],
								filenames[i],
								filesizes[i],
								prefix,
								urls[i]));
				});
			} else {
				console.error(messages.cant_handle_this_file(prefix + filenames[i]));
			}
		}
		
		return p.then(() => {
			return files;
		})
	})
}

function loop_until_files_are_downloaded(needle_options, download_dir) {
	aria2.removeAllListeners("onDownloadComplete")
		.removeAllListeners("onDownloadError")
		.removeAllListeners("onDownloadStop");
	//promise1 tries adds tasks to Aria2, and then do it again after 5 minutes.
	let promise1 = download_list_of_files(needle_options, download_dir)
		.then(GID => {
			console.log(messages.print_GID(GID));
			console.log(messages.retry_in_N_minutes(retry_minutes));
			return timer();
		})
		.catch(error => {
			if (error == messages.sharepoint_link_download_finished ||
				error == messages.sharepoint_token_expired ||
				error.toString().startsWith(messages.unknown_status_code_startswith)) //throw error immediately.
				throw(error);
			else {
				console.error(error);
				console.log(messages.retry_in_N_minutes(retry_minutes));
				return timer();
			}
		})
	//promise2 listens to Aria2 events, return/reject when anything happens.
	//It's so that Aria2 can add new tasks when completes/errors happen, without waiting for N minutes.
	let promise2 = new Promise((resolve, reject) => {
			aria2.on("onDownloadComplete", ([GID]) => {
				GID = GID.gid;
				console.log(messages.Aria2_download_complete(GID_filepath_dictionary[GID]));
				delete GID_filepath_dictionary[GID];
				resolve(GID);
			});
			aria2.on("onDownloadError", ([GID]) => {
				GID = GID.gid;
				console.error(messages.Aria2_download_errored(GID_filepath_dictionary[GID]));
				delete GID_filepath_dictionary[GID];
				reject(GID);
			});
			aria2.on("onDownloadStop", ([GID]) => {
				GID = GID.gid;
				console.error(messages.Aria2_download_stopped(GID_filepath_dictionary[GID]));
				delete GID_filepath_dictionary[GID];
				reject(GID);
			});
		})
	//Race between promise1 and promise2
	return Promise.race([promise1, promise2])
	.then(() =>
		loop_until_files_are_downloaded(needle_options)
	)
	.catch((error) => {
		if (error == messages.sharepoint_link_download_finished ||
			error == messages.sharepoint_token_expired ||
			error.toString().startsWith(messages.unknown_status_code_startswith)) //throw error immediately.
			throw(error);
		else {
			return loop_until_files_are_downloaded(needle_options);
		}
	})
}

function download_list_of_files(needle_options) {
	let files_that_is_downloading_in_aria2 = [];
	
	return aria2.call("tellActive", ["gid","files"])
	.then(tasks => {
		//Get list of ongoing Aria2 download tasks.
		if (tasks.length>0)
			for (const task of tasks)
				for (const file of task["files"]) {
					files_that_is_downloading_in_aria2.push(file["path"])
					//When you start the script, there might be tasks already downloading in Aria2.
					//So we register these GIDs/paths
					GID_filepath_dictionary[task["gid"]] = file["path"];
				}
		if (tasks.length >= aria2_max_concurrent_downloads)
			throw(messages.Aria2_tasks_full);
		else
			return;
	})
	.then(() => {
		//Filter filelist to get the files that's going to be added to Aria2.
		const filesNeeded = aria2_max_concurrent_downloads - files_that_is_downloading_in_aria2.length;
		const list_of_files_to_parse = [];
		for (let i=0; i<Math.min(files_to_be_downloaded.length, filesNeeded); i++) {
			if (files_that_is_downloading_in_aria2.includes(files_to_be_downloaded[i].path)) {
				continue;
			} else if (!fs.existsSync(files_to_be_downloaded[i].path) || fs.existsSync(files_to_be_downloaded[i].path + ".aria2")) {
				list_of_files_to_parse.push(files_to_be_downloaded[i]);
			} else if (fs.statSync(files_to_be_downloaded[i].path).size < 1024 && files_to_be_downloaded[i].filesize != fs.statSync(files_to_be_downloaded[i].path).size) {
				//If you got throttled, your downloaded file size is 576B, you need to delete the file and re-add tasks to Aria2.
				//Luckily I am doing this for you.
				
				//After tests, it turns out if you use this script it's "very unlikely" to get 576B files.
				//I'm not removing this section because... I'm not sure 576B files won't happen with this script.
				fs.unlinkSync(files_to_be_downloaded[i].path);
				list_of_files_to_parse.push(files_to_be_downloaded[i]);
			} else {
				//Remove already downloaded files from files_to_be_downloaded
				files_to_be_downloaded.splice(i--, 1);
			}
		}
			
		if (list_of_files_to_parse.length == 0)
			if (files_that_is_downloading_in_aria2.length == 0)
				throw(messages.sharepoint_link_download_finished);
			else
				throw(messages.no_tasks_to_add);
		return list_of_files_to_parse;
	})
	.then(files => {
		//Add tasks to Aria2.
		const promises = [];
		let header="Cookie: "; //This header will be passed to Aria2.
		for (const i in needle_options.cookies)
			header += i + "=" + needle_options.cookies[i] + "; ";
		files.forEach(file => {
			let promise = needle("get", file.APIurl, needle_options)
				.then(response => {
					if (response.body["webUrl"]) {
						return needle("head",response.body["webUrl"], needle_options)
							.then(resp => {
								//Notes about status codes:
								//sharepoint returns 429 when throttleing your service.
								//sharepoint returns 401 when token is expired.
								//sharepoint returns 200 when everything is okay.
								
								//The original implementation is to pass response.body["@content.downloadUrl"],
								//but that link is often throttled by sharepoint server.
								//I don't know why ["webUrl"] is less likely to be throttled...but it is
								if (resp.statusCode == 200) {
									if (response.body["file"]["mimeType"].startsWith("application/vnd")) {
										//application/vnd.... is mimetype for documents that will probably open onedrive APPs
										//(which will not let us download the file.)
										//We want to use response.body["@content.downloadUrl"] in this case.
										return aria2.call("addUri",
											[response.body["@content.downloadUrl"]],
											{dir: path.dirname(file.path)}
										)
									} else {
										return aria2.call("addUri",
											[response.body["webUrl"]],
											{dir: path.dirname(file.path),
											 header: header}
										)
									}
									
								} else if (resp.statusCode == 401) {
									throw(messages.sharepoint_token_expired);
								} else if (resp.statusCode == 429 || resp.statusCode == 503) {
									throw(messages.sharepoint_throttled);
								} else {
									throw(messages.unknown_status_code(resp.statusCode));
								}
							})
					} else
						throw(messages.sharepoint_token_expired);
				})
				.then(GID => {
					console.log(messages.Aria2_tasks_added(file.prefix + file.filename));
					GID_filepath_dictionary[GID] = file.prefix + file.filename;
					return GID;
				})
			promises.push(promise);
		})
		//return Promise.all(promises);
		if (promises.length == 0) {
			return timer(20);
		} else {
			return Promise.all(promises);
		}
	})
}

function linkinfo(url_, dir_, pwd_ = null) {
	return {
		url: url_,
		dir: dir_,
		pwd: pwd_
	}
}

function fileInfo(filepath_, filename_, filesize_, prefix_, url_) {
	return {
		path: filepath_,
		filename: filename_,
		filesize: filesize_,
		prefix: prefix_,
		APIurl: url_
	}
}

function timer(seconds=retry_minutes*60) {
	return new Promise(resolve => {
		setTimeout(resolve, seconds*1000)
	});
}

function sanitize_Unicode_encoded_string(str) {
	//Code taken from https://github.com/CodeDotJS/unicodechar-string/blob/master/index.js
	//Licensed under MIT.
	//I copied the code because it's unneccessary to add a npm dependancy
	return str.replace(/\\u[\dA-Fa-f]{4}/g, match => {
		return String.fromCharCode(parseInt(match.replace(/\\u/g, ''), 16));
	});
}

//messages consists of message strings in this program
//It's helpful for fixing typos, translations or adding colors to messages if you want.
//I used var here so I can declare this at the buttom of this .js file
const messages = {
	"Aria2_IP_not_specified": `<error> You need to put "Aria2_IP" field in "${aria2_connection_info_filepath}".`,
	"Aria2_port_not_specified": `<error> You need to put "Aria2_port" field in "${aria2_connection_info_filepath}".`,
	"Aria2_token_not_specified": `<warning> There is no "Aria2_token" field in "${aria2_connection_info_filepath}", can possibly cause connection issues.`,
	"Aria2_connection_opening": "<info> opening aria2 connection.",
	"Aria2_connection_opened": "<info> aria2 connection opened.",
	"Invalid_linklist": "<error> Invalid linklist.",
	"invalid_linklist_syntax": line => `<error> line ${line} in linklist file has invalid linklist syntax, skipped.`,
	"not_absolute_path": path => `<error> "${path}" is not absolute file path, which is require to make this program work.`,
	"no_valid_sharepoint_link": "<error> No valid sharepoint link is found.",
	"found_N_sharepoint_links": num => `<info> found ${num} sharepoint links.`,
	"program_finished": "<info> No more sharepoint link to download, program ended.",
	"folder_is_password_protected": "<info> Getting password protected folder.",
	"password_incorrect": url => `<error> Given password is incorrect for password protected folder (${url})`,
	"password_incorrect_startswith": `<error> Given password is incorrect for password protected folder`, //whatever "password_incorrect" startswith
	"password_not_provided": url => `<error> Password in not provided for password protected folder (${url})`,
	"password_not_provided_startswith": `<error> Password in not provided for password protected folder`, //whatever "password_not_provided" startswith
	"redirect_url_not_found": "<error> Redirect url not found.",
	"sharepoint_link_download_finished": "<info> All files in this sharepoint link downloaded.",
	"retry_in_N_minutes": minutes => `<info> Retry in ${minutes} minutes.`,
	"got_N_file_informations": num => `<info> Got file information of ${num} files.`,
	"getting_sharepoint_file_list": "<info> Getting file list of the sharepoint folder, it might take a while...",
	"folder_traverse_failed": (path, error) => `<error> Folder traversal failed in "${path}".\n${error}`,
	"cant_handle_this_file": path => `<error> "${path}" is neither a file nor a folder, I can't handle it.`,
	"print_GID": GID => `<info> Aria2 GID: ${GID}`,
	"sharepoint_token_expired": "<error> Sharepoint token expired.",
	"sharepoint_throttled": "<error> Throttled by sharepoint server.",
	"unknown_status_code": code => `<error> Unknown status code returned by sharepoint server: ${code}.`,
	"unknown_status_code_startswith": "<error> Unknown status code returned by sharepoint server", //whatever "unknown_status_code" startswith,
	"Aria2_download_complete": path => `<info> Download complete - "${path}"`,
	"Aria2_download_errored": path => `<error> Download errored - "${path}"`, //I chose to use the word "errored", maybe there is a more proper way to express this.
	"Aria2_download_stopped": path => `<error> Download stopped - "${path}"`,
	"Aria2_tasks_full": "<info> Aria2 has reached maximum active tasks.",
	"no_tasks_to_add": "<info> No new files to add to Aria2, wait for current downloads to complete.",
	"Aria2_tasks_added": path => `<info> "${path}" is added to Aria2 download task.`,
	"Aria2_concurrent_downloads": num => `<info> This program will run up to ${num} download tasks at the same time.`
}