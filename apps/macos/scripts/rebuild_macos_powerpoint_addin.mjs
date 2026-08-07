process.stderr.write(
  "Disabled permanently: PowerPoint VBA must be compiled inside Office for Mac. " +
    "Use a macro-enabled PPTM, import the reviewed .bas/.cls modules in the VBE, " +
    "run Debug > Compile VBAProject, save, fully quit PowerPoint, and then run " +
    "package_macos_offline_addins.mjs to inject the compiled vbaProject.bin into " +
    "a verified PPAM shell.\n",
);
process.exitCode = 2;
