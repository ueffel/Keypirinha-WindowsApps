from .lib import helper
import keypirinha as kp
import keypirinha_util as kpu
import subprocess
import os
import glob
import time
import asyncio
import traceback


class WindowsApps(kp.Plugin):
    """Lists Universal Windows Apps (formerly Metro Apps) in Keypirinha for launching
    """

    DEFAULT_ITEM_LABEL = "Windows App:"
    DEFAULT_SHOW_MISC_APPS = False
    DEFAULT_PREFERRED_CONTRAST = "black"

    def __init__(self):
        """Default constructor and initializing internal attributes
        """
        super().__init__()
        self._item_label = self.DEFAULT_ITEM_LABEL

    def _get_icon(self, name, icon_path):
        """Create a list of possible logo files to show as icon for a window app
        """
        base_path = os.path.splitext(icon_path)
        logos = []
        logos.extend(glob.glob(icon_path))
        logos.extend(glob.glob("{}.scale-*{}".format(base_path[0], base_path[1])))
        logos.extend(glob.glob("{}/scale-*/{}{}".format(os.path.dirname(base_path[0]),
                                                        os.path.basename(base_path[0]),
                                                        base_path[1])))
        logos.extend(glob.glob("{}.contrast-*{}".format(base_path[0], base_path[1])))
        logos.extend(glob.glob("{}/contrast-*/{}{}".format(os.path.dirname(base_path[0]),
                                                           os.path.basename(base_path[0]),
                                                           base_path[1])))
        logos.extend(glob.glob("{}/contrast-*/{}.contrast-*{}".format(os.path.dirname(base_path[0]),
                                                                      os.path.basename(base_path[0]),
                                                                      base_path[1])))
        logos.extend(glob.glob("{}/contrast-*/{}.scale-*{}".format(os.path.dirname(base_path[0]),
                                                                   os.path.basename(base_path[0]),
                                                                   base_path[1])))
        logos.extend(glob.glob("{}/contrast-*/scale-*/{}{}".format(os.path.dirname(base_path[0]),
                                                                   os.path.basename(base_path[0]),
                                                                   base_path[1])))
        logos.extend(glob.glob("{}/scale-*/{}.contrast-*{}".format(os.path.dirname(base_path[0]),
                                                                   os.path.basename(base_path[0]),
                                                                   base_path[1])))
        logos.extend(glob.glob("{}/scale-*/contrast-*/{}{}".format(os.path.dirname(base_path[0]),
                                                                   os.path.basename(base_path[0]),
                                                                   base_path[1])))

        logos_preferred = [logo for logo in logos if "contrast-{}".format(self._preferred_contrast) in logo]
        if logos_preferred:
            logos = logos_preferred

        self.dbg(name)
        for logo in logos:
            self.dbg("{}".format(logo))

        if logos:
            cached_logos = self._copy_files(name, logos)
            handle = self.load_icon(cached_logos)
            return handle

    def _copy_files(self, name, logos):
        """Copies the logos to the package cache and returns a list of keypirinha resource strings
        """
        cached_logos = []
        for logo in logos:
            out_dir = os.path.join(self.get_package_cache_path(True), name)
            if not os.path.isdir(out_dir):
                os.mkdir(out_dir)
            out_path = os.path.join(out_dir, os.path.basename(logo))
            if not os.path.isfile(out_path):
                with open(logo, "rb") as in_file, \
                        open(out_path, "wb") as out_file:
                    out_file.write(in_file.read())
            cached_logos.append("cache://{}/{}/{}".format(self.package_full_name(),
                                                          name,
                                                          os.path.basename(logo)))
        return cached_logos

    def _read_config(self):
        """Reads the default action from the config
        """
        self.dbg("Reading config")
        settings = self.load_settings()

        self._debug = settings.get_bool("debug", "main", False)

        self._item_label = settings.get("item_label", "main", self.DEFAULT_ITEM_LABEL)
        self.dbg("item_label =", self._item_label)

        self._show_misc_apps = settings.get_bool("show_misc_apps", "main", self.DEFAULT_SHOW_MISC_APPS)
        self.dbg("show_misc_apps =", self._show_misc_apps)

        self._preferred_contrast = settings.get_enum("preferred_contrast",
                                                     "main",
                                                     self.DEFAULT_PREFERRED_CONTRAST,
                                                     ["black", "white"])
        self.dbg("preferred_contrast =", self._preferred_contrast)

    def on_start(self):
        """Reads the config
        """
        self._read_config()

    def on_events(self, flags):
        """Reloads the package config when its changed
        """
        if flags & kp.Events.PACKCONFIG:
            self._read_config()

    def on_catalog(self):
        """Catalogs items for keypirinha

        Calls powershell to get a list of windows app packages with their properties
        and creates catalog items for keypirinha.
        """
        start_time = time.time()
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        output, err = subprocess.Popen(["powershell.exe",
                                        "mode con cols=512; Get-AppxPackage"],
                                       stdout=subprocess.PIPE,
                                       universal_newlines=True,
                                       shell=False,
                                       startupinfo=startupinfo).communicate()

        tasks = []
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        # packages a separated by a double newline within the output
        for package in output.strip().split("\n\n"):
            # collect all the properties into a dict
            props = {}
            for line in package.splitlines():
                idx = line.index(":")
                key = line[:idx].strip()
                value = line[idx + 1:].strip()
                props[key] = value

            tasks.append(asyncio.ensure_future(self._create_catalog_item(props)))

        job = asyncio.gather(*tasks, return_exceptions=True)
        loop.run_until_complete(job)
        loop.close()

        catalog = []
        for task in tasks:
            exc = task.exception()
            if exc:
                self.warn(exc)
                self.dbg("".join(traceback.format_exception(exc.__class__, exc, exc.__traceback__)))
                continue
            items = task.result()
            if items:
                catalog.extend(items)

        self.set_catalog(catalog)
        elapsed = time.time() - start_time
        self.info("Cataloged {} items in {:0.1f} seconds".format(len(catalog), elapsed))

    async def _create_catalog_item(self, props):
        try:
            package = helper.AppXPackage(props)
            # some packages can be just libraries with no executable application
            # only take packages which have a application
            apps = await package.apps()
            catalog_items = []
            if apps:
                self.dbg(package.InstallLocation)
                for app in apps:
                    if not app.misc_app or self._show_misc_apps:
                        catalog_items.append(self.create_item(
                            category=kp.ItemCategory.CMDLINE,
                            label="{} {}".format(self._item_label, app.display_name).strip(),
                            short_desc=app.description
                            if app.description else app.display_name,
                            target=app.execution,
                            args_hint=kp.ItemArgsHint.FORBIDDEN,
                            hit_hint=kp.ItemHitHint.NOARGS,
                            icon_handle=self._get_icon(package.Name, app.icon_path)
                        ))
            return catalog_items
        except Exception as exc:
            if "Name" in props:
                raise Exception("Error while creating catalog item for '{0}'".format(props["Name"])) from exc
            else:
                raise Exception("Error while creating catalog item for {0}".format(props)) from exc

    def on_execute(self, item, action):
        """Starts the windows app
        """
        self.dbg("Executing:", item.target())
        kpu.shell_execute("explorer.exe", item.target())
