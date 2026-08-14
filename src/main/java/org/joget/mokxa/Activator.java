package org.joget.mokxa;

import java.util.ArrayList;
import java.util.Collection;
import org.osgi.framework.BundleActivator;
import org.osgi.framework.BundleContext;
import org.osgi.framework.ServiceRegistration;

public class Activator implements BundleActivator {

    protected Collection<ServiceRegistration> registrationList;

    public void start(BundleContext context) {
        registrationList = new ArrayList<ServiceRegistration>();

        //Register plugin here
        registrationList.add(context.registerService(OutlookEmailListBinder.class.getName(), new OutlookEmailListBinder(), null));
        registrationList.add(context.registerService(OutlookInboxListBinder.class.getName(), new OutlookInboxListBinder(), null));
        registrationList.add(context.registerService(OutlookSentListBinder.class.getName(), new OutlookSentListBinder(), null));

    }

    public void stop(BundleContext context) {
        for (ServiceRegistration registration : registrationList) {
            registration.unregister();
        }
    }
}