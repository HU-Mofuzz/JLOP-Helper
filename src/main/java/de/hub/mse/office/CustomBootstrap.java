package de.hub.mse.office;

import com.sun.star.bridge.UnoUrlResolver;
import com.sun.star.bridge.XUnoUrlResolver;
import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.connection.NoConnectException;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import helper.Lo;

import java.io.*;
import java.util.Map;
import java.util.Random;

public class CustomBootstrap {

    public static final XComponentContext bootstrap(String[] var0) throws BootstrapException {
        XComponentContext context = null;

        try {
            XComponentContext componentContext = Bootstrap.createInitialComponentContext((Map)null);
            if (componentContext == null) {
                throw new BootstrapException("no local component context!");
            } else {
                File officeExecutable = new File("/usr/lib/libreoffice/program/soffice");
                String pipeName = "uno" + (new Random().nextLong() & Long.MAX_VALUE);
                String[] executeArgs = new String[var0.length + 2];
                executeArgs[0] = officeExecutable.getPath();
                executeArgs[1] = "--accept=pipe,name=" + pipeName + ";urp;";
                System.arraycopy(var0, 0, executeArgs, 2, var0.length);
                Process executesProcess = Runtime.getRuntime().exec(executeArgs);
                pipe(executesProcess.getInputStream(), System.out, "CO> ");
                pipe(executesProcess.getErrorStream(), System.err, "CE> ");
                XMultiComponentFactory multiComponentFactory = componentContext.getServiceManager();
                if (multiComponentFactory == null) {
                    throw new BootstrapException("no initial service manager!");
                } else {
                    XUnoUrlResolver urlResolver = UnoUrlResolver.create(componentContext);
                    String unoUrl = "uno:pipe,name=" + pipeName + ";urp;StarOffice.ComponentContext";
                    int retries = 0;

                    while(true) {
                        try {
                            Object xInterface = urlResolver.resolve(unoUrl);
                            context = UnoRuntime.queryInterface(XComponentContext.class, xInterface);
                            if (context == null) {
                                throw new BootstrapException("no component context!");
                            }

                            return context;
                        } catch (NoConnectException var13) {
                            if (retries++ == 300) {
                                throw new BootstrapException(var13);
                            }

                            Thread.sleep(100L);
                        }
                    }
                }
            }
        } catch (BootstrapException var14) {
            throw var14;
        } catch (RuntimeException var15) {
            throw var15;
        } catch (Exception var16) {
            throw new BootstrapException(var16);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        } catch (java.lang.Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static void pipe(final InputStream var0, final PrintStream var1, final String var2) {
        (new Thread("Pipe: " + var2) {
            public void run() {
                try {
                    BufferedReader var1x = new BufferedReader(new InputStreamReader(var0, "UTF-8"));

                    while(true) {
                        String var2x = var1x.readLine();
                        if (var2x == null) {
                            break;
                        }

                        var1.println(var2 + var2x);
                    }
                } catch (UnsupportedEncodingException var3) {
                    var3.printStackTrace(System.err);
                } catch (IOException var4) {
                    var4.printStackTrace(System.err);
                }

            }
        }).start();
    }

    public static XComponentLoader getLoaderFromContext(XComponentContext context) {
        if(context == null) {
            return null;
        }
        // get the remote office service manager
        var mcFactory = context.getServiceManager();
        if (mcFactory == null) {
            System.out.println("Office Service Manager is unavailable");
            System.exit(1);
        }

        // desktop service handles application windows and documents
        XDesktop xDesktop = null;
        try {
            xDesktop = Lo.qi(XDesktop.class, mcFactory.createInstanceWithContext("com.sun.star.frame.Desktop", context));
        } catch (Exception e) {
            return null;
        }
        if (xDesktop == null) {
            System.out.println("Could not create a desktop service");
            System.exit(1);
        }

        // XComponentLoader provides ability to load components
        return Lo.qi(XComponentLoader.class, xDesktop);
    }

    public static XComponentLoader bootstrap() {
        try {
            return getLoaderFromContext(bootstrap(Bootstrap.getDefaultOptions()));
        } catch (BootstrapException e) {
            return null;
        }
    }
}
